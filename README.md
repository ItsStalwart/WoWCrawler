# WoWCrawler
HTML Crawler for retrieving equipment information from wowprogress.

Add desired characters to track by copying wowprogress profile link to charactersToLoad list. 

i.e: charactersToLoad = new List<string>() {"https://www.wowprogress.com/character/us/azralon/Stalwart"};
Please note this is not yet optimized for performance so it will use up a lot of RAM if you add too many characters to the list.

File output directory may be edited freely on line 166 ("FileInfo excelFile = new FileInfo(@"YOUR DIRECTORY GOES HERE"+  fileName + ".xlsx");").

No data is saved until all data for all characters in the list has been crawled.

Please note this may stop working at any point if wowprogress applies changes to their website structure.
