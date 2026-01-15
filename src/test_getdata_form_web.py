import asyncio
from get_data_from_web import StudentScraper

async def main():
    scraper = StudentScraper(url="http://localhost:5500/Edit-TblTranscripts-Table.html")
    await scraper.open_web()

    # ดึงข้อมูลจากหน้าแรก
    data1 = await scraper.get_data()
    for x in data1:
        print(x['index'], x['id'], x['name'])

    # เปลี่ยนไปหน้าใหม่
    #await scraper.change_page("http://localhost:5500/Edit-TblTranscripts-Table2.html")
    #data2 = await scraper.get_data()
    #print("Page 2:", data2)

    await scraper.close_web()

asyncio.run(main())