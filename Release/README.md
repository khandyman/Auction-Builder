Auction Builder

This program assists in making "auction" macros in the TAKP EverQuest client, as implemented on the Project Quarm emulated server.  The program requires use of the Zeal add-on package created by Salty for use on Quarm.


Installation
Unzip Auction-Builder.zip into any directory.  Ensure that all files in the zip remain in the same location.  Then launch Auction-Builder.exe.


Program Description
 - Auction builder reads a zeal outputfile of a character's inventory and assembles a list of all the items present.  
 - Then it performs an HTML scrape of all these item's URLs on www.eqtunnelauctions.com to find their auction history.  This history is averaged to calculate a price for every item in the list. (Note: this function is a little slow, due to the poor performance of Python's request package. In tests, a list of 42 items took approximately 1.5 minutes)
 - Finally, this data is converted into a format the EverQuest client will recognize as item links and written into the character's .ini file where it will be available in game as a macro button.


Basic Use
 - To use Auction Builder, click the Import button, then wait while the program assembles a list of items and prices, which it will print out.  
 - The list may then be examined and the calculated prices can be adjusted as desired.
 - Additionally, if an item should be excluded from this and any future macros, check the exclude box for that item.  
 - Finally, click Save and the list will be written to the .ini file.

Settings Description
 - Hotkey Starting Location: specify the page and button where auction macros should begin. Each macro will store up to 30 items.  If more than 30 items are available for sale, multiple macros will be created, incrementing the button by 1 each time.
 - Auctions to Count: indicate the desired number of auctions from www.eqtunnelauctions.com to use in calculating an average price.  If an item has less than the number specified, the program will use as many as it can find.  If an item has no auction data, the price will be listed as 'unknown'.
 - Outputfile Path: this is the path to a character's Zeal outputfile.  Simply click the text field to change the file path.
 - Mule Ini Path: this is the path to a character's EQ .ini file.  Simply click the text field to change the file path.
 - Item Exclusions: if items are present in the character's inventory that should not be sold, they can be marked for exclusion.  Items in this list will be ignored.