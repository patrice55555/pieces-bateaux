const PRODUCT_SITEMAP_URL = "https://parts.princess.co.uk/product-sitemap.xml";
const DEFAULT_BATCH_SIZE = 20;
const MAX_BATCH_SIZE = 20;
const PRODUCT_CONCURRENCY = 4;

const CSV_COLUMNS = [
  "url",
  "category",
  "name",
  "description",
  "specifications",
  "model_number",
  "price_gbp_numeric",
  "stock",
  "delivery_delay",
];

export default {
  async fetch(request) {
    const url = new URL(request.url);

    if (url.pathname === "/") {
      return jsonResponse({
        message: "Princess Parts scraper",
        endpoints: {
          manifest: "/manifest",
          csv: "/scrape.csv?offset=0&limit=20",
          json: "/scrape.json?offset=0&limit=20",
        },
        note: "Le scraping complet doit etre fait par lots pour rester compatible avec les limites Cloudflare Workers.",
      });
    }

    if (url.pathname === "/manifest") {
      const productUrls = await getProductUrls();
      return jsonResponse({
        total: productUrls.length,
        defaultBatchSize: DEFAULT_BATCH_SIZE,
        maxBatchSize: MAX_BATCH_SIZE,
        sample: productUrls.slice(0, 5),
      });
    }

    if (url.pathname === "/all.csv") {
      return new Response(
        "Utilisez /scrape.csv avec offset et limit pour iterer par lots.",
        { status: 400, headers: { "content-type": "text/plain; charset=UTF-8" } },
      );
    }

    if (url.pathname === "/scrape.csv" || url.pathname === "/scrape.json") {
      const productUrls = await getProductUrls();
      const offset = clampInteger(url.searchParams.get("offset"), 0, productUrls.length);
      const requestedLimit = clampInteger(url.searchParams.get("limit"), DEFAULT_BATCH_SIZE, MAX_BATCH_SIZE);
      const batchUrls = productUrls.slice(offset, offset + requestedLimit);
      const products = await mapWithConcurrency(batchUrls, PRODUCT_CONCURRENCY, scrapeProductPage);
      const rows = products.map(stripInternalFields);
      const failedCount = products.filter((product) => product._error).length;
      const nextOffset = offset + batchUrls.length;
      const hasMore = nextOffset < productUrls.length;

      if (url.pathname === "/scrape.json") {
        return jsonResponse({
          total: productUrls.length,
          offset,
          limit: requestedLimit,
          returned: rows.length,
          failed: failedCount,
          nextOffset: hasMore ? nextOffset : null,
          hasMore,
          products: rows,
        });
      }

      return new Response(toCsv(rows, CSV_COLUMNS), {
        headers: {
          "content-type": "text/csv; charset=UTF-8",
          "content-disposition": `attachment; filename="princess-parts-${offset}-${offset + batchUrls.length - 1}.csv"`,
          "x-total-count": String(productUrls.length),
          "x-next-offset": hasMore ? String(nextOffset) : "",
          "x-failed-count": String(failedCount),
        },
      });
    }

    return new Response("Not found", { status: 404 });
  },
};

async function getProductUrls() {
  const response = await fetch(PRODUCT_SITEMAP_URL, {
    headers: {
      "user-agent": "Cloudflare-Workers-Princess-Parts-Scraper/1.0",
    },
  });

  if (!response.ok) {
    throw new Error(`Impossible de recuperer le sitemap produit (${response.status})`);
  }

  const xml = await response.text();

  return [...xml.matchAll(/<loc>(.*?)<\/loc>/gsi)]
    .map((match) => match[1].trim())
    .filter((entry) => entry.includes("/product/"));
}

async function scrapeProductPage(productUrl) {
  try {
    const response = await fetch(productUrl, {
      headers: {
        "user-agent": "Cloudflare-Workers-Princess-Parts-Scraper/1.0",
      },
    });

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }

    const html = await response.text();
    const parsed = await parseProductHtml(html);

    return {
      url: productUrl,
      category: parsed.categories.join(" | "),
      name: parsed.name,
      description: parsed.description.join("\n"),
      specifications: parsed.specifications,
      model_number: parsed.modelNumber,
      price_gbp_numeric: parsed.priceGbpNumeric,
      stock: parsed.stock,
      delivery_delay: parsed.delivery,
      _error: "",
    };
  } catch (error) {
    return {
      url: productUrl,
      category: "",
      name: "",
      description: "",
      specifications: "",
      model_number: "",
      price_gbp_numeric: "0.00",
      stock: "0",
      delivery_delay: "",
      _error: error instanceof Error ? error.message : String(error),
    };
  }
}

async function parseProductHtml(html) {
  const product = {
    name: "",
    priceGbp: "",
    priceGbpNumeric: "0.00",
    stock: "0",
    modelNumber: "",
    categories: [],
    description: [],
    accordionBlocks: [],
    specifications: "",
    delivery: "",
  };

  await new HTMLRewriter()
    .on("h1.product_title", firstText((value) => {
      product.name = value;
    }))
    .on("p.price .woocommerce-Price-amount bdi", firstText((value) => {
      product.priceGbp = extractPrice(value);
      product.priceGbpNumeric = normalizePrice(value);
    }))
    .on("p.stock", firstText((value) => {
      product.stock = normalizeStock(value);
    }))
    .on(".product_meta .sku_wrapper .sku", firstText((value) => {
      product.modelNumber = value;
    }))
    .on(".product_meta .posted_in a", manyText((value) => {
      if (!product.categories.includes(value)) {
        product.categories.push(value);
      }
    }))
    .on("div.woocommerce-product-details__short-description p", manyText((value) => {
      product.description.push(value);
    }))
    .on("#accordion li", manyText((value) => {
      product.accordionBlocks.push(value);
    }))
    .transform(new Response(html))
    .text();

  for (const block of product.accordionBlocks) {
    if (/^specifications?\b/i.test(block)) {
      product.specifications = stripSectionHeading(block, /^specifications?\b[:\s-]*/i);
      continue;
    }

    if (/^delivery\b/i.test(block)) {
      product.delivery = stripSectionHeading(block, /^delivery\b[:\s-]*/i);
    }
  }

  return product;
}

function firstText(onValue) {
  let hasValue = false;
  return textCollector((value) => {
    if (!hasValue && value) {
      hasValue = true;
      onValue(value);
    }
  });
}

function manyText(onValue) {
  return textCollector(onValue);
}

function textCollector(onValue) {
  return {
    buffers: [],
    element(element) {
      const current = { text: "" };
      this.buffers.push(current);
      element.onEndTag(() => {
        const finished = this.buffers.pop();
        const value = normalizeText(finished?.text ?? "");
        if (value) {
          onValue(value);
        }
      });
    },
    text(text) {
      if (this.buffers.length > 0) {
        this.buffers[this.buffers.length - 1].text += text.text;
      }
    },
  };
}

function stripSectionHeading(value, pattern) {
  return normalizeText(value.replace(pattern, ""));
}

function extractPrice(value) {
  const match = value.match(/([0-9]+(?:[.,][0-9]{2})?)/);
  return match ? match[1].replace(",", ".") : "";
}

function normalizePrice(value) {
  if (!value) {
    return "0.00";
  }

  const match = value.replace(/,/g, "").match(/(\d+(?:\.\d+)?)/);
  if (!match) {
    return "0.00";
  }

  const amount = Number.parseFloat(match[1]);
  return Number.isFinite(amount) ? amount.toFixed(2) : "0.00";
}

function normalizeStock(value) {
  if (!value) {
    return "0";
  }

  const match = value.match(/(\d+)/);
  return match ? match[1] : "0";
}

function stripInternalFields(product) {
  return {
    url: product.url,
    category: product.category,
    name: product.name,
    description: product.description,
    specifications: product.specifications,
    model_number: product.model_number,
    price_gbp_numeric: product.price_gbp_numeric,
    stock: product.stock,
    delivery_delay: product.delivery_delay,
  };
}

function toCsv(rows, columns) {
  const header = columns.join(",");
  const body = rows
    .map((row) => columns.map((column) => escapeCsvValue(row[column] ?? "")).join(","))
    .join("\n");

  return `${header}\n${body}`;
}

function escapeCsvValue(value) {
  const normalized = String(value).replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  if (!/[",\n]/.test(normalized)) {
    return normalized;
  }

  return `"${normalized.replace(/"/g, '""')}"`;
}

function normalizeText(value) {
  return value.replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
}

function clampInteger(value, fallback, max) {
  const parsed = Number.parseInt(value ?? "", 10);
  if (!Number.isFinite(parsed) || parsed < 0) {
    return fallback;
  }

  return Math.min(parsed, max);
}

async function mapWithConcurrency(items, concurrency, mapper) {
  const results = new Array(items.length);
  let nextIndex = 0;

  async function worker() {
    while (nextIndex < items.length) {
      const currentIndex = nextIndex;
      nextIndex += 1;
      results[currentIndex] = await mapper(items[currentIndex], currentIndex);
    }
  }

  const workers = Array.from({ length: Math.min(concurrency, items.length) }, () => worker());
  await Promise.all(workers);
  return results;
}

function jsonResponse(data, status = 200) {
  return new Response(JSON.stringify(data, null, 2), {
    status,
    headers: {
      "content-type": "application/json; charset=UTF-8",
    },
  });
}