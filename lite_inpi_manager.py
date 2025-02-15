import os
import scrapy
import pandas as pd
from scrapy.crawler import CrawlerRunner
from twisted.internet import reactor
from twisted.internet.defer import DeferredList

class InpiSpider(scrapy.Spider):
    name = "inpi"
    start_urls = ["https://busca.inpi.gov.br/pePI/servlet/LoginController?action=login"]
    custom_settings = {
        'DOWNLOAD_DELAY': 5,
        'AUTOTHROTTLE_ENABLED': True,
    }
    
    def __init__(self, patent_number=None, *args, **kwargs):
        super(InpiSpider, self).__init__(*args, **kwargs)
        self.patent_number = patent_number
        self.data = []

    def parse(self, response):
        next_page_link = response.css('area[data-mce-href="menu-servicos/patente"]::attr(href)').get()
        yield response.follow(next_page_link, callback=self.parse_next_page)

    def parse_next_page(self, response):
        yield scrapy.FormRequest.from_response(
            response,
            formdata={"NumPedido": self.patent_number},
            callback=self.parse_patent_details,
        )

    def parse_patent_details(self, response):
        next_page_link = response.css('a[class="visitado"]::attr(href)').get()
        yield response.follow(next_page_link, callback=self.extract_search)

    def extract_search(self, response):
        elementos_texto = response.css('a')
        for elemento in elementos_texto:
            texto = elemento.css('a.normal[href="javascript:void(0)"]::text').get()
            if texto:
                self.data.append(texto.strip())

class InpiManager:
    def __init__(self):
        self.dir_path = os.path.dirname(os.path.realpath(__file__))
        self.excel_file = 'input_example.xlsx'  # Excel file in same directory

    def check_status(self, codes):
        nvigente_codes = ['8.12', '11.1.1', '11.2', '11.4', '11.6', '11.11', '11.21', 
                         '18.3', '21.1', '21.2', '21.7', '11.20', '9.2.4', '111', '112', '113']
        
        for code in codes:
            if any(nv_code in code for nv_code in nvigente_codes):
                return "NÃO VIGENTE"
        return "VIGENTE"

    def process_patents(self):
        try:
            # Read Excel file
            df = pd.read_excel(self.excel_file)
            patent_numbers = df['Nº DA PROTEÇÃO'].dropna().tolist()

            # Setup crawler
            runner = CrawlerRunner()
            deferreds = []
            results = {}

            # Process each patent
            for patent in patent_numbers:
                spider = InpiSpider(patent_number=patent)
                deferred = runner.crawl(spider)
                deferreds.append(deferred)
                results[patent] = spider.data

            # Wait for all spiders to finish
            dlist = DeferredList(deferreds)
            dlist.addBoth(lambda _: reactor.stop())
            reactor.run()

            # Update Excel with results
            for patent, codes in results.items():
                mask = df['Nº DA PROTEÇÃO'] == patent
                df.loc[mask, 'STATUS'] = self.check_status(codes)

            # Save updated Excel file
            df.to_excel(self.excel_file, index=False)
            print("Excel file updated successfully!")

        except Exception as e:
            print(f"Error: {e}")

def main():
    manager = InpiManager()
    manager.process_patents()

if __name__ == "__main__":
    main()
