import os
import scrapy
import pandas as pd
from scrapy.crawler import CrawlerProcess
from scrapy.crawler import CrawlerRunner
from twisted.internet import reactor
from twisted.internet.defer import DeferredList

class StatusCodes:
    NVIGENTE = {
        '8.12': '8.12 - ARQ DEFINITIVO - FALTA DE PGT',
        '11.1.1': '11.1.1 - ARQ DEFINITIVO - ANTERIORIDADE',
        # ... (rest of nvigente)
    }
    
    VIGENTE = {
        '9.1': '9.1 - DEFERIMENTO',
        '16.1': '16.1 - CONCESSÃO DE CARTA PATENTE',
        # ... (rest of vigente)
    }
    
    ANALISE_SUB = {
        '16.1': '16.1 - CONCESSÃO DE CARTA PATENTE',
        # ... (rest of analise_sub)
    }

class InpiSpiderProtection(scrapy.Spider):
    name = "inpi_protection"
    start_urls = ["https://busca.inpi.gov.br/pePI/servlet/LoginController?action=login"]
    custom_settings = {
        'DOWNLOAD_DELAY': 5,
        'AUTOTHROTTLE_ENABLED': True,
        'AUTOTHROTTLE_START_DELAY': 1.0,
        'AUTOTHROTTLE_MAX_DELAY': 60.0,
        'AUTOTHROTTLE_TARGET_CONCURRENCY': 10.0,
    }
    
    def __init__(self, protection_number=None, *args, **kwargs):
        super(InpiSpiderProtection, self).__init__(*args, **kwargs)
        self.protection_number = protection_number
        self.data = []

    def parse(self, response):
        next_page_link = response.css('area[data-mce-href="menu-servicos/patente"]::attr(href)').get()
        yield response.follow(next_page_link, callback=self.parse_next_page)

    def parse_next_page(self, response):
        yield scrapy.FormRequest.from_response(
            response,
            formdata={"NumPedido": self.protection_number},
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
        self.excel_file = os.path.join(self.dir_path, 'input_example.xlsx')

    def read_protection_numbers(self):
        """Read BR patent numbers from Excel file"""
        try:
            df = pd.read_excel(self.excel_file)
            return df['Nº DA PROTEÇÃO'].dropna().tolist()
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            return []

    def analyze_protection(self, exig):
        """Analyze protection status and return results"""
        exig_nvig = []
        exig_ansu = []
        exig_vig = []
        status = 0

        for item in exig:
            if item in StatusCodes.NVIGENTE:
                exig_nvig.append(StatusCodes.NVIGENTE[item])
            if item in StatusCodes.VIGENTE:
                exig_vig.append(StatusCodes.VIGENTE[item])
            if item in StatusCodes.ANALISE_SUB:
                exig_ansu.append(StatusCodes.ANALISE_SUB[item])

        if exig_nvig:
            status = 1
        else:
            if '9.1' not in StatusCodes.VIGENTE:
                if '120' not in exig_ansu:
                    exig_ansu.append('- PARECER TÉCNICO EMITIDO -')

        if '203 - EXAME TÉCNICO SOLICITADO' not in exig_vig:
            exig_ansu.append('- EXAME TÉCNICO AUSENTE!!! -')

        despacho = exig_nvig + exig_vig

        return {
            'status': 'NÃO VIGENTE' if status else 'VIGENTE',
            'despacho': '; '.join(despacho),
            'analise': '; '.join(exig_ansu)
        }

    def process_protections(self, protection_numbers):
        runner = CrawlerRunner()
        deferreds = []
        results = {}

        for prot in protection_numbers:
            spider = InpiSpiderProtection(protection_number=prot)
            deferred = runner.crawl(spider)
            deferreds.append(deferred)
            results[prot] = spider.data

        dlist = DeferredList(deferreds)
        dlist.addBoth(lambda _: reactor.stop())
        reactor.run()

        return results

    def update_excel(self, results):
        """Update Excel file with results"""
        try:
            df = pd.read_excel(self.excel_file)
            
            for prot, data in results.items():
                if prot in df['Nº DA PROTEÇÃO'].values:
                    analysis = self.analyze_protection(data)
                    idx = df[df['Nº DA PROTEÇÃO'] == prot].index[0]
                    
                    df.at[idx, 'STATUS'] = analysis['status']
                    df.at[idx, 'DESPACHO'] = analysis['despacho']
                    df.at[idx, 'ANÁLISE SUBSTANTIVA'] = analysis['analise']
                    
                    print(f"Updated protection {prot}")

            df.to_excel(self.excel_file, index=False)
            print("Excel file updated successfully")
            
        except Exception as e:
            print(f"Error updating Excel file: {e}")

def main():
    manager = InpiManager()
    
    # Read protection numbers from Excel
    protection_numbers = manager.read_protection_numbers()
    if not protection_numbers:
        print("No protection numbers found in Excel file")
        return

    print(f"Processing {len(protection_numbers)} protection numbers...")
    
    # Process each protection
    results = manager.process_protections(protection_numbers)
    
    # Update Excel file with results
    manager.update_excel(results)

if __name__ == "__main__":
    main()
