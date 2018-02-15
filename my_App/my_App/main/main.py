from mainScraper import mainScraper

def main():
	scraper = mainScraper()

	for i in range(1,10):
		if(i != 4):   # needed since para04 is missing          
			scraper.dataExtraction("../inputFiles/para0" + str(i) + ".pdf")  #path strings to the pdfs

	scraper.writeToExcel("../outputFiles/BeneFix Small Group Plans upload template.xlsx") #path string to the excel file
	print("The scraping is done!! Check the excel file")

if __name__ == '__main__':
	main()