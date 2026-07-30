import fs from "node:fs/promises";
import path from "node:path";
import readline from "node:readline/promises";
import process from "node:process";
import { chromium } from "playwright";

function samePublicationHost(candidateUrl, hostname) {
  try {
    const candidateHost = new URL(candidateUrl).hostname
      .toLowerCase()
      .replace(/^www\./, "");
    return candidateHost === hostname || candidateHost.endsWith(`.${hostname}`);
  } catch {
    return false;
  }
}

function isLikelyArticleUrl(candidateUrl, hostname) {
  if (!samePublicationHost(candidateUrl, hostname)) return false;
  try {
    const pathname = new URL(candidateUrl).pathname;
    return !/\/(?:ads?|places|business|classifieds?|marketplace|coupons?|directories|category|categories|forms?|tag|author|search|account|login|subscribe)(?:\/|$)/i
      .test(pathname);
  } catch {
    return false;
  }
}

function cleanText(value, maximumLength) {
  return String(value || "")
    .replace(/\u0000/g, "")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim()
    .slice(0, maximumLength);
}

function looksLikeAccessChallenge(title, text, url) {
  const sample = `${title}\n${text}\n${url}`.toLowerCase();
  const hardChallenge = [
    "verify you are human",
    "checking your browser",
    "access denied",
    "enable javascript and cookies",
    "captcha",
    "challenge-platform",
  ].some((phrase) => sample.includes(phrase));
  const loginGate = [
    "sign in to continue",
    "log in to continue",
    "subscriber-only",
    "subscription required",
  ].some((phrase) => sample.includes(phrase));
  return hardChallenge || (loginGate && String(text || "").trim().length < 1_500);
}

function decodeXml(value) {
  return String(value || "")
    .replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, "$1")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}

function extractXmlEntries(xml, blockName) {
  const entries = [];
  const blockPattern = new RegExp(
    `<${blockName}\\b[^>]*>([\\s\\S]*?)<\\/${blockName}>`,
    "gi",
  );
  for (const match of String(xml || "").matchAll(blockPattern)) {
    const block = match[1];
    const loc = decodeXml(
      block.match(/<loc\b[^>]*>([\s\S]*?)<\/loc>/i)?.[1] || "",
    ).trim();
    const lastmod = decodeXml(
      block.match(/<(?:news:publication_date|lastmod)\b[^>]*>([\s\S]*?)<\/(?:news:publication_date|lastmod)>/i)?.[1] || "",
    ).trim();
    const title = decodeXml(
      block.match(/<(?:news:)?title\b[^>]*>([\s\S]*?)<\/(?:news:)?title>/i)?.[1] || "",
    ).trim();
    if (loc) entries.push({ loc, lastmod, title });
  }
  return entries;
}

function developmentKeyword(value) {
  return /\b(demolish|demolition|redevelop|redevelopment|development|apartments?|housing|residential|mixed-use|retail|restaurant|shopping center|grocery|construction|site plan)\b/i
    .test(String(value || "").replace(/[-_/]+/g, " "));
}

async function fetchBrowserText(page, url) {
  const response = await page.request.get(url, {
    timeout: 8_000,
    failOnStatusCode: false,
  });
  if (!response.ok()) return "";
  return (await response.text()).slice(0, 8_000_000);
}

async function discoverFromSitemaps(page, source, window, config) {
  const origin = new URL(source.sourceUrl).origin;
  const sitemapSeeds = [
    `${origin}/sitemap.xml`,
    `${origin}/sitemap_index.xml`,
    `${origin}/news-sitemap.xml`,
    `${origin}/sitemap-news.xml`,
  ];
  const sitemapUrls = [];
  const articleEntries = [];
  for (const seed of sitemapSeeds) {
    try {
      const xml = await fetchBrowserText(page, seed);
      if (!xml) continue;
      articleEntries.push(...extractXmlEntries(xml, "url"));
      sitemapUrls.push(...extractXmlEntries(xml, "sitemap").map((item) => item.loc));
    } catch {
      // Publications frequently omit one or more conventional sitemap URLs.
    }
  }
  const prioritizedSitemaps = [...new Set(sitemapUrls)]
    .sort((a, b) =>
      Number(/\b(news|post|article|story)\b/i.test(b)) -
      Number(/\b(news|post|article|story)\b/i.test(a)))
    .slice(0, Number(config.browser?.maximumSitemapsPerSite || 20));
  for (const sitemapUrl of prioritizedSitemaps) {
    try {
      const xml = await fetchBrowserText(page, sitemapUrl);
      articleEntries.push(...extractXmlEntries(xml, "url"));
    } catch {
      // Continue with other sitemap files.
    }
  }
  const datedStart = `${window.start}T00:00:00`;
  const datedEnd = `${window.end}T23:59:59`;
  return articleEntries
    .filter((item) => isLikelyArticleUrl(item.loc, source.hostname))
    .filter((item) =>
      !/\/(?:ads?|places|business|classifieds?|marketplace|coupons?|directories|tag|author|search)(?:\/|$)/i
        .test(new URL(item.loc).pathname))
    .map((item) => {
      const dateValue = item.lastmod ? new Date(item.lastmod) : null;
      const inWindow = dateValue && !Number.isNaN(dateValue.getTime())
        ? dateValue >= new Date(datedStart) && dateValue <= new Date(datedEnd)
        : false;
      return {
        ...item,
        score:
          (developmentKeyword(`${item.title} ${item.loc}`) ? 2 : 0) +
          (inWindow ? 2 : 0),
      };
    })
    .filter((item) =>
      item.score >= 4 &&
      developmentKeyword(`${item.title} ${item.loc}`))
    .sort((a, b) => b.score - a.score || b.lastmod.localeCompare(a.lastmod))
    .map((item) => item.loc);
}

async function waitForLogin(page, publicationName) {
  const terminal = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
  try {
    await terminal.question(
      `Login or complete the verification for ${publicationName} in the open browser, then press Enter here to continue... `,
    );
    await page.reload({ waitUntil: "domcontentloaded", timeout: 90_000 });
  } finally {
    terminal.close();
  }
}

async function extractPageEvidence(page, source, config) {
  const maximumArticleCharacters = Number(
    config.browser?.maximumArticleCharacters || 45_000,
  );
  return page.evaluate(({ maximumArticleCharacters, hostname }) => {
    const absoluteUrl = (value) => {
      try {
        return new URL(value, document.baseURI).href;
      } catch {
        return "";
      }
    };
    const meta = (selector, attribute = "content") =>
      document.querySelector(selector)?.getAttribute(attribute)?.trim() || "";
    const jsonLd = [];
    for (const node of document.querySelectorAll('script[type="application/ld+json"]')) {
      try {
        const parsed = JSON.parse(node.textContent || "null");
        jsonLd.push(...(Array.isArray(parsed) ? parsed : [parsed]));
      } catch {
        // Invalid JSON-LD is common and does not prevent extraction.
      }
    }
    const flattened = jsonLd.flatMap((item) => item?.["@graph"] || item || []);
    const articleData = flattened.find((item) => {
      const type = item?.["@type"];
      const types = Array.isArray(type) ? type : [type];
      return types.some((entry) =>
        ["Article", "NewsArticle", "ReportageNewsArticle", "BlogPosting"]
          .includes(entry));
    }) || {};
    const primary =
      document.querySelector("article") ||
      document.querySelector("main") ||
      document.querySelector('[role="main"]') ||
      document.body;
    const pdfLinks = [...document.querySelectorAll("a[href]")]
      .map((anchor) => ({
        url: absoluteUrl(anchor.getAttribute("href")),
        label: (anchor.innerText || anchor.getAttribute("aria-label") || "").trim(),
      }))
      .filter((item) =>
        item.url &&
        (/\.pdf(?:$|[?#])/i.test(item.url) ||
          /\b(plan|plans|site plan|floor plan|elevation|design review|specific plan)\b/i
            .test(item.label)))
      .slice(0, 50);
    const canonical = absoluteUrl(
      meta('link[rel="canonical"]', "href") || location.href,
    );
    const text = (primary?.innerText || document.body?.innerText || "")
      .replace(/\u0000/g, "")
      .replace(/[ \t]+\n/g, "\n")
      .replace(/\n{3,}/g, "\n\n")
      .trim()
      .slice(0, maximumArticleCharacters);
    return {
      requested_url: location.href,
      canonical_url: canonical,
      publication_hostname: hostname,
      title:
        meta('meta[property="og:title"]') ||
        articleData.headline ||
        document.title ||
        "",
      publication_date:
        meta('meta[property="article:published_time"]') ||
        meta('meta[name="date"]') ||
        meta('time[datetime]', "datetime") ||
        articleData.datePublished ||
        "",
      author:
        meta('meta[name="author"]') ||
        articleData.author?.name ||
        "",
      description:
        meta('meta[property="og:description"]') ||
        meta('meta[name="description"]') ||
        articleData.description ||
        "",
      text,
      pdf_links: pdfLinks,
    };
  }, { maximumArticleCharacters, hostname: source.hostname });
}

async function waitForPublicationSearchResults(page, config) {
  const timeout = Number(
    config.browser?.searchResultsWaitMilliseconds || 12_000,
  );
  await page.waitForTimeout(
    Number(config.browser?.searchMinimumWaitMilliseconds || 5_000),
  );
  await page.waitForLoadState("networkidle", {
    timeout: Math.min(timeout, 6_000),
  }).catch(() => {});
  await page.waitForFunction(
    () => [...document.querySelectorAll("a[href]")].some((anchor) =>
      /\b(demolish|demolition|redevelop|development|apartments?|housing|residential|mixed-use|retail|restaurant|shopping center|grocery|construction|site plan)\b/i
        .test(`${anchor.innerText || ""} ${anchor.getAttribute("href") || ""}`
          .replace(/[-_/]+/g, " "))),
    null,
    { timeout },
  ).catch(() => {});
  await page.evaluate(() =>
    window.scrollTo({ top: document.body.scrollHeight, behavior: "instant" }))
    .catch(() => {});
  await page.waitForTimeout(1_500);
}

async function discoverCandidateUrls(page, source, window, config) {
  const navigationTimeout = Number(
    config.browser?.navigationTimeoutMilliseconds || 30_000,
  );
  const urls = await discoverFromSitemaps(
    page,
    source,
    window,
    config,
  ).catch(() => []);
  console.log(`    Sitemap discovery found ${urls.length} candidate URL(s).`);

  const origin = new URL(source.sourceUrl).origin;
  for (const query of ["development", "housing apartments", "retail restaurant"]) {
    if (urls.length >= Number(config.browser?.maximumCandidatesPerSite || 10)) {
      break;
    }
    try {
      await page.goto(
        `${origin}/search/?q=${encodeURIComponent(query)}`,
        { waitUntil: "domcontentloaded", timeout: navigationTimeout },
      );
      console.log(`    Waiting for internal search results: ${query}`);
      await waitForPublicationSearchResults(page, config);
      const siteResults = await page.locator("a[href]").evaluateAll((anchors) =>
        anchors
          .filter((anchor) =>
            /\b(demolish|demolition|redevelop|development|apartments?|housing|residential|mixed-use|retail|restaurant|shopping center|grocery|construction|site plan)\b/i
              .test(`${anchor.innerText || ""} ${anchor.getAttribute("href") || ""}`
                .replace(/[-_/]+/g, " ")))
          .map((anchor) => anchor.href)
          .filter(Boolean));
      for (const url of siteResults) {
        if (
          isLikelyArticleUrl(url, source.hostname) &&
          !urls.includes(url)
        ) {
          urls.push(url);
        }
      }
    } catch {
      // Many sites use a different search route; other discovery methods remain.
    }
  }
  console.log(`    Publication search found ${urls.length} total candidate URL(s).`);

  try {
    await page.goto(source.sourceUrl, {
      waitUntil: "domcontentloaded",
      timeout: navigationTimeout,
    });
    await page.waitForTimeout(1_000);
    const homepageLinks = await page.locator("a[href]").evaluateAll((anchors) =>
      anchors
        .filter((anchor) =>
          /\b(demolition|development|apartments?|housing|mixed-use|retail|restaurant|shopping center|redevelopment)\b/i
            .test(anchor.innerText || anchor.getAttribute("aria-label") || ""))
        .map((anchor) => anchor.href)
        .filter(Boolean));
    for (const url of homepageLinks) {
      if (isLikelyArticleUrl(url, source.hostname) && !urls.includes(url)) {
        urls.push(url);
      }
    }
  } catch (error) {
    console.warn(`    Homepage discovery warning: ${error.message}`);
  }
  console.log(`    Homepage discovery found ${urls.length} total candidate URL(s).`);

  return urls.slice(0, Number(config.browser?.maximumCandidatesPerSite || 10));
}

export async function startBrowserSession(scriptDirectory, options, config) {
  const profileDirectory = path.resolve(
    scriptDirectory,
    config.browser?.profileDirectory || "browser-profile",
  );
  await fs.mkdir(profileDirectory, { recursive: true });
  const channel = config.browser?.channel || "msedge";
  const context = await chromium.launchPersistentContext(profileDirectory, {
    channel,
    headless: Boolean(options.headless),
    chromiumSandbox: true,
    ignoreDefaultArgs: ["--no-sandbox"],
    viewport: { width: 1440, height: 1000 },
    acceptDownloads: true,
    locale: "en-US",
    args: [
      "--no-first-run",
      "--no-default-browser-check",
      "--disable-features=msEdgeFirstRunExperience",
    ],
  });
  context.setDefaultTimeout(45_000);
  return context;
}

export async function prepareSiteLogins(context, sources) {
  const page = context.pages()[0] || await context.newPage();
  for (let index = 0; index < sources.length; index += 1) {
    const source = sources[index];
    console.log(
      `[${index + 1}/${sources.length}] Login setup: ${source.aliases.join("; ")} (${source.sourceUrl})`,
    );
    try {
      await page.goto(source.sourceUrl, {
        waitUntil: "domcontentloaded",
        timeout: 90_000,
      });
    } catch (error) {
      console.warn(`  Page load warning: ${error.message}`);
    }
    await waitForLogin(page, source.aliases.join("; "));
  }
  await page.goto("about:blank").catch(() => {});
}

export async function collectSourceEvidence(
  context,
  source,
  window,
  config,
  options,
) {
  const page = context.pages()[0] || await context.newPage();
  const evidence = [];
  let challengeCount = 0;
  const navigationTimeout = Number(
    config.browser?.navigationTimeoutMilliseconds || 30_000,
  );
  try {
    console.log(`    Opening publication homepage: ${source.sourceUrl}`);
    try {
      await page.goto(source.sourceUrl, {
        waitUntil: "domcontentloaded",
        timeout: navigationTimeout,
      });
      await page.bringToFront();
      await page.waitForTimeout(
        Number(config.browser?.pageSettleMilliseconds || 1_500),
      );
      const homepageText = await page.locator("body").innerText().catch(() => "");
      if (
        looksLikeAccessChallenge(
          await page.title(),
          homepageText,
          page.url(),
        )
      ) {
        challengeCount += 1;
        console.warn(
          "    Access challenge detected on the homepage; skipping this publication.",
        );
        return {
          evidence: [],
          audit: {
            status: "blocked",
            notes:
              "Browser detected a login, Continue, CAPTCHA, or access challenge on the publication homepage. The unattended run skipped this publication.",
          },
        };
      }
    } catch (error) {
      console.warn(`    Homepage load warning: ${error.message}`);
    }
    console.log("    Discovering candidate articles...");
    const candidates = await discoverCandidateUrls(
      page,
      source,
      window,
      config,
    );
    for (let index = 0; index < candidates.length; index += 1) {
      const candidate = candidates[index];
      console.log(`    Browser article ${index + 1}/${candidates.length}: ${candidate}`);
      const articlePage = await context.newPage();
      try {
        await articlePage.bringToFront();
        await articlePage.goto(candidate, {
          waitUntil: "domcontentloaded",
          timeout: navigationTimeout,
        });
        await articlePage.waitForTimeout(
          Number(config.browser?.pageSettleMilliseconds || 1_500),
        );
        let extracted = await extractPageEvidence(
          articlePage,
          source,
          config,
        );
        if (looksLikeAccessChallenge(
          extracted.title,
          extracted.text,
          extracted.requested_url,
        )) {
          challengeCount += 1;
          console.warn(
            `    Access challenge detected; skipping article: ${candidate}`,
          );
          continue;
        }
        if (cleanText(extracted.text, 1_000).length >= 200) {
          evidence.push(extracted);
        }
      } catch (error) {
        console.warn(`    Could not extract ${candidate}: ${error.message}`);
      } finally {
        await articlePage.close().catch(() => {});
      }
    }
    const status = evidence.length
      ? "reviewed"
      : challengeCount
        ? "blocked"
        : "no_results";
    return {
      evidence,
      audit: {
        status,
        notes: evidence.length
          ? `Browser extracted ${evidence.length} candidate article(s); ${challengeCount} access challenge(s) encountered.`
          : challengeCount
            ? `Browser encountered ${challengeCount} login/CAPTCHA/access challenge(s) and extracted no usable article text.`
            : "Browser search found no usable candidate article text in the selected date window.",
      },
    };
  } finally {
    await page.goto("about:blank").catch(() => {});
  }
}
