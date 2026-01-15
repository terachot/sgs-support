# get_data_from_web.py
from playwright.async_api import async_playwright

class StudentScraper:
    def __init__(self, url=None):
        self.url = url
        self.playwright = None
        self.browser = None
        self.page = None

    async def open_web(self):
        """เปิด browser และเข้าไปยังหน้าแรก"""
        self.playwright = await async_playwright().start()
        self.browser = await self.playwright.chromium.launch(channel="chrome", headless=False)
        self.page = await self.browser.new_page()
        if self.url:
            await self.page.goto(self.url)

    async def change_page(self, url: str):
        """เปลี่ยนไปหน้าใหม่"""
        if not self.page:
            raise RuntimeError("Browser ยังไม่ถูกเปิดด้วย open_web()")
        await self.page.goto(url)
    
    async def login_web(self, username, password):
        """เข้าระบบเพื่อใช้งาน"""
        input_username = "input[name='ctl00$PageContent$UserName']" #ช่องกรอก username
        input_password = "input[name='ctl00$PageContent$Password']" #ช่องกรอก password
        button_login = "a[id='ctl00_PageContent_OKButton__Button']"  #ปุ่ม login
        home_workpage = "http://127.0.0.1:5500/Show-TblSchoolInfo.html" #หน้าการใช้งานแรก

        print(username)
        print(password)

        await self.page.locator(input_username).fill(username)
        await self.page.locator(input_password).fill(password)
        await self.page.locator(button_login).click()

        if self.page.url == home_workpage:
            return True
        else:
            return False

    async def get_data(self):
        """ดึงข้อมูลจากหน้าเว็บปัจจุบัน"""
        rows = self.page.locator('td.tre tbody tr')
        count = await rows.count()
        data = []

        for i in range(count):
            row = rows.nth(i)
            cells = await row.locator("td").all_text_contents()
            if not cells:
                continue
            student_id = cells[3] if len(cells) > 3 else None
            student_name = cells[4] if len(cells) > 4 else None
            data.append({
                "index": i,
                "id": student_id,
                "name": student_name
            })
        return data #type list

    async def get_url(self):
        """หน้าเว็บปัจจุบัน"""
        return self.page.url

    async def logout_web(self):
        """ออกจากระบบ"""
        button_logout = "a[id='ctl00__PageHeader__SignIn']" #ปุ่ม logout
        
        await self.page.locator(button_logout).click()

    async def close_web(self):
        """ปิด browser และ playwright"""
        if self.browser:
            await self.browser.close()
        if self.playwright:
            await self.playwright.stop()
        return True