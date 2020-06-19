from mechanicalsoup import StatefulBrowser
from threading import Thread
import xlrd
import time

# from win10toast import ToastNotifier
from plyer import notification


with open('config.txt', 'r') as file:
    lines = file.read().split('\n')
    for line in lines:
        if line.lower().startswith('site:'):
            main_url = line[5:].strip()

            if main_url.endswith('/'):
                main_url = main_url[:-1]

        if line.lower().startswith('filename:'):
            filename = line[9:].strip()

        if line.lower().startswith('login:'):
            login = line[6:].strip()
        if line.lower().startswith('password:'):
            password = line[9:].strip()


class CompanyData:
    companies = []
    not_added_companies = []
    cant_be_added_companies = []

    @staticmethod
    def read_data(filename):
        wb = xlrd.open_workbook(filename)
        ws = wb.sheet_by_index(0)

        data = []
        for i in range(1, ws.nrows):
           data.append(ws.row_values(i))

        return data

    @staticmethod
    def fill_data(filename='Data.xlsx'):
        for company_data in CompanyData.read_data(filename):
            if company_data[0].strip() == '':
                continue

            company = CompanyData(company_data)
            CompanyData.companies.append(company)

            if company.name not in added_companies:
                CompanyData.not_added_companies.append(company)


    def __init__(self, data):
        self.name = data[0].strip()
        self.category = data[1].strip()
        self.tag = data[2].strip()
        self.logo_link = data[3].strip() if len(data[3].strip()) > 3 else ''
        self.website = data[4].strip() if len(data[4].strip()) > 3 else ''
        self.email = data[5].strip() if len(data[5].strip()) > 3 else ''
        self.linkedin = data[6].strip() if len(data[6].strip()) > 3 else ''
        self.facebook = data[7].strip() if len(data[7].strip()) > 3 else ''
        self.twitter = data[8].strip() if len(data[8].strip()) > 3 else ''
        self.short_desc = data[9].strip()
        self.desc = data[10].strip()
        self.products = data[11].strip()


with open('Added companies.txt', 'r') as file:
    added_companies = file.read().split('\n')


CompanyData.fill_data(filename)


def create_browser():
    browser = StatefulBrowser()

    browser.open(f'{main_url}/wp-login.php')

    browser.select_form('form[name="loginform"]')
    browser['log'] = login
    browser['pwd'] = password

    browser.submit_selected()

    return browser


def add_companies(companies):
    global number_of_finished_threads

    browser = create_browser()
    for company in companies:
        if company.name in added_companies:
            continue

        try_counter = 0
        while True:
            if try_counter > 0:
                CompanyData.cant_be_added_companies.append(company)
                break

            try:
                try_counter += 1

                browser.open(f'{main_url}/wp-admin/post-new.php?post_type=company')

                browser.select_form('form[name="post"]')

                browser['post_title'] = company.name
                browser['content'] = company.desc
                browser['excerpt'] = company.short_desc
                browser['acf[field_5ec410bcf22d5]'] = company.website
                browser['acf[field_5ec410d2f22d6]'] = company.email
                browser['acf[field_5ec41118f22d8]'] = company.linkedin
                browser['acf[field_5ec410f6f22d7]'] = company.facebook
                browser['acf[field_5ec52f4834bd5]'] = company.twitter
                browser['acf[field_5ec4e629c23f6]'] = company.products

                browser.submit_selected()

                print(company.name)

                break
            except Exception as error:
                raise error
                print("Error!-----------------------------------------------------------------------------------")
                print(repr(error))
                print()

        raise Exception('dfgkporegkoek')

    number_of_finished_threads += 1


# number_of_threads = (len(CompanyData.not_added_companies) - 1) // 10 + 1
number_of_finished_threads = 0
# for i in range(number_of_threads):
#     thread = Thread(target=add_companies, args=(CompanyData.not_added_companies[10 * i: 10 * i + 10],))
#     thread.start()
    # thread.join()

# while number_of_finished_threads < number_of_threads:
#     time.sleep(5)

# time.sleep(60)


add_companies(CompanyData.not_added_companies)


added_companies = [company.name for company in CompanyData.companies if company not in CompanyData.cant_be_added_companies]
with open('Added companies.txt', 'w') as file:
    file.write('\n'.join(added_companies))


add_companies(CompanyData.cant_be_added_companies)

cant_be_added_companies = [company.name for company in CompanyData.cant_be_added_companies if company not in CompanyData.added_companies]
with open('Companies that can\'t be added.txt', 'w') as file:
    file.write('\n'.join(cant_be_added_companies))


# notifier = ToastNotifier()
# notifier.show_toast("Program finished!", "Enjoy")
notification.notify("Program finished!", "Enjoy")