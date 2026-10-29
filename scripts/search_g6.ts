import ZAI from 'z-ai-web-dev-sdk';
import { writeFileSync } from 'fs';

const companies = [
  "ООО ЗНВ лакокрасочные материалы бренды краски",
  "ООО РЭЙФОР лакокрасочные материалы краски",
  "ООО СТРОЙТОРГ ГАЛАКТИКА лакокрасочные краски бренды",
  "ООО ТД Хоздвор Лакра Кубанские Краски ассортимент",
  "ООО УНИВЕРСАЛ-ЮГ лакокрасочные материалы",
  "ООО Электротехническая компания лакокрасочные краски",
  "Паритет ООО 7726425616 лакокрасочные краски",
  "Снегирев Владимир ИП лакокрасочные краски",
  "СтройБаза-14 ТК ООО лакокрасочные краски",
  "Сыргабаев Куштарбек ИП лакокрасочные краски",
  "ИП Попов А.В. лакокрасочные краски",
  "Хабитуев Александр Валерьевич ИП лакокрасочные",
  "ЭЛИТА-СЕРВИС ЦТО ООО лакокрасочные краски"
];

async function main() {
  const zai = await ZAI.create();
  
  for (let i = 0; i < companies.length; i++) {
    const query = companies[i];
    console.log(`Searching ${i+1}/${companies.length}: ${query}`);
    
    try {
      const results = await zai.functions.invoke('web_search', {
        query: query,
        num: 5
      });
      
      writeFileSync(`/tmp/search_g6_${i}.json`, JSON.stringify(results, null, 2));
      console.log(`  Found ${results.length} results`);
    } catch (error: any) {
      console.error(`  Error: ${error.message}`);
      writeFileSync(`/tmp/search_g6_${i}.json`, JSON.stringify({ error: error.message }));
    }
    
    // Wait 5 seconds between searches
    if (i < companies.length - 1) {
      console.log('  Waiting 5s...');
      await new Promise(resolve => setTimeout(resolve, 5000));
    }
  }
  
  console.log('Done!');
}

main().catch(console.error);
