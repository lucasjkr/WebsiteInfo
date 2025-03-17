import requests, json
from requests.packages.urllib3.exceptions import InsecureRequestWarning
from openpyxl import Workbook
import bs4

requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

class WebChain:

    def __init__(self):
        self.headers = {
            'User-Agent': "Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:136.0) Gecko/20100101 Firefox/136.0",
        }

    def redirect_chain(self, url):
        # dictionary to insert current query results into
        data = {}
        data['request_url'] = url
        data['status_code'] = str("")
        data['title'] = str("")
        data['final_url'] = str("")
        data['chain'] = list()

        # if supplied url starts with http, request the URL as is
        if url.startswith('http'):
            response = requests.get(url, allow_redirects=True, verify=False, timeout=1, headers=self.headers)
        else:
            # otherwise, request that url with HTTPS and if that fails, with HTTP
            try:
                url_target = f"https://{url}"
                response = requests.get(url_target, allow_redirects=True, verify=False, timeout=1, headers=self.headers)
            except Exception:
                try:
                    url_target = f"http://{url}"
                    response = requests.get(url_target, allow_redirects=True, verify=False, timeout=1, headers=self.headers)
                except Exception:
                    # If both attempts failed, return the URL with a status of 0, so we know the URL was attempted
                    data['status_code'] = 0
                    data['chain'] = ""
                    return data

        data['status_code'] = response.status_code

        # check whether HTTP response included redirects
        history = response.history
        if history:
            for resp in history:
                # get each step in the responses redirect history
                data['chain'].append(resp.url)
            # append the final URL to the end of the redirect chain
            data['chain'].append(response.url)
        else:
            # if there were no redirects, then the response chain is just the requested URL
            data['chain'].append(response.url)

        data['chain'] = json.dumps(data['chain'], indent=4)
        data['final_url'] = response.url
        data['title'] = self.get_page_title(response.content)
        return data

    def get_page_title(self, content):
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

        # format the resulting spreadsheet
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

        # Freeze top rows
        worksheet.freeze_panes = 'A2'
        workbook.remove(workbook['Sheet'])
        workbook.save("output.xlsx")


    def main(self, path_to_url_file):
        # create empty list to insert all results into
        result = []

        # open input text file and break into lines (1 URL per line)
        # Url's can just be fqdn or entire fqdn with protocol and may include an optional port number
        with open(path_to_url_file, 'r') as file:
            urls = [line.rstrip() for line in file]
            for url in urls:
                # if URL starts with "#" then skip
                if url.startswith('#'):
                    print(f"Skipping {url}")
                    continue
                else:
                    # print(url)
                    # chain = self.redirect_chain(url)
                    # result.append( chain )
                    # query website info and append to result set
                    result.append( self.redirect_chain(url) )
        # write results to a text file
        self.write_to_excel(result)

if __name__ == "__main__":
    test = WebChain()
    test.main('input.txt')
    print("done!")