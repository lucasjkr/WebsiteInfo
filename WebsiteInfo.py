import requests, json
from requests.packages.urllib3.exceptions import InsecureRequestWarning
from openpyxl import Workbook
import bs4

requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

class WebChain:

    def __init__(self):
        self.useragent = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:136.0) Gecko/20100101 Firefox/136.0"
        self.headers = {
            'User-Agent': self.useragent,
        }
        pass

    def redirect_chain(self, url):
        data = {}
        data['request_url'] = url
        data['status_code'] = str("")
        data['title'] = str("")
        data['final_url'] = str("")
        data['chain'] = list()

        if url.startswith('http'):
            response = requests.get(url, allow_redirects=True, verify=False, timeout=1, headers=self.headers)
        else:
            try:
                url_target = f"https://{url}"
                response = requests.get(url_target, allow_redirects=True, verify=False, timeout=1, headers=self.headers)
            except Exception:
                try:
                    url_target = f"http://{url}"
                    response = requests.get(url_target, allow_redirects=True, verify=False, timeout=1, headers=self.headers)
                except Exception:
                    data['status_code'] = 0
                    data['chain'] = ""
                    return data
        data['status_code'] = response.status_code

        history = response.history
        if history:
            for resp in history:
                data['chain'].append(resp.url)
            data['chain'].append(response.url)
        else:
            data['chain'].append(response.url)

        data['chain'] = json.dumps(data['chain'], indent=4)

        data['final_url'] = response.url
        data['title'] = self.get_title(response.content)
        return data

    def get_title(self, content):
        soup = bs4.BeautifulSoup(content, 'html.parser')
        title = soup.find('title')
        if title:
            return title.string
        else:
            return ""

    def write_to_excel(self, results):
        workbook = Workbook()

        for result in results:
            if result is None:
                continue
            else:
                # If sheet named 'Data' does not exist, create it
                if 'Data' not in workbook:
                    worksheet = workbook.create_sheet(title='Data')
                    worksheet.append([key for key in result])

                if result is not None:
                    # now insert the row of values that need to be inserted (action happens whether header row was created or not
                    worksheet = workbook['Data']
                    worksheet.append(list(result.values()))

        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2) * 1.0
            worksheet.column_dimensions[column_letter].width = adjusted_width

        worksheet.freeze_panes = 'A2'
        workbook.remove(workbook['Sheet'])
        workbook.save("output.xlsx")


    def main(self, path_to_url_file):
        result = []
        with open(path_to_url_file, 'r') as file:
            urls = [line.rstrip() for line in file]
            for url in urls:
                if url.startswith('#'):
                    print(f"Skipping {url}")
                    continue
                else:
                    print(url)
                    chain = self.redirect_chain(url)
                    result.append( chain )
        self.write_to_excel(result)

if __name__ == "__main__":
    test = WebChain()
    test.main('input.txt')
    print("done!")