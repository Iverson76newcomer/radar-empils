import ZAI from '/home/z/.bun/install/global/node_modules/z-ai-web-dev-sdk/dist/index.js';
import { writeFileSync, readFileSync } from 'fs';

const queries = [
  { idx: 4, query: "КОРВОН Душанбе краски ЛКМ бренды Турция Иран" },
  { idx: 15, query: "Касторама Россия ЛКМ краски бренды ассортимент Dulux Tikkurila" },
  { idx: 20, query: "МАСТАК-Плюс Пермь краски ЛКМ бренды ассортимент" },
  { idx: 23, query: "ОБИ Россия ЛКМ краски бренды ассортимент Dulux Caparol" },
  { idx: 24, query: "Басар ООО краски ЛКМ бренды ассортимент" },
  { idx: 25, query: "ГИГАНТСТРОЙ ООО краски ЛКМ бренды" },
  { idx: 14, query: "Фаткуллин ИП краски ЛКМ бренды" },
  { idx: 19, query: "Воронков ИП краски ЛКМ бренды Макаров" },
  { idx: 10, query: "Баткомплект ООО ЛКМ краски бренды" },
  { idx: 18, query: "ЛКМ комплект ООО Тверь краски бренды" },
];

async function main() {
  const zai = await ZAI.create();

  for (const { idx, query } of queries) {
    let retries = 5;
    while (retries > 0) {
      try {
        console.log(`[${new Date().toISOString()}] Searching idx ${idx}: ${query.substring(0, 50)}...`);
        const res = await zai.functions.invoke('web_search', { query, num: 5 });
        writeFileSync(`/tmp/search_g5_${idx}.json`, JSON.stringify(res, null, 2));
        console.log(`  ✅ Got ${res.length} results`);
        break;
      } catch (e) {
        retries--;
        if (e.message.includes('429')) {
          const wait = 30;
          console.log(`  ⏳ Rate limited, waiting ${wait}s... (${retries} retries left)`);
          await new Promise(r => setTimeout(r, wait * 1000));
        } else {
          console.log(`  ❌ Error: ${e.message.substring(0, 80)}`);
          break;
        }
      }
    }
    // Wait between successful requests
    await new Promise(r => setTimeout(r, 5000));
  }

  console.log("\n=== ALL DONE ===");
}

main().catch(console.error);
