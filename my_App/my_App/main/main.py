from mainScraper import mainScraper

def main():
	scraper = mainScraper()

	for i in range(1,10):
		if(i != 4):
			scraper.dataExtraction("../inputFiles/para0" + str(i) + ".pdf")

	scraper.writeToExcel("../outputFiles/BeneFix Small Group Plans upload template.xlsx")

main()