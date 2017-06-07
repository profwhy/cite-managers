# cite-managers
EBSCOHost XML Parser Excel macro

This Excel Macro takes XML-formatted citation output from EBSCOHost, and extracts certain elements into the worksheet. It is used to create a neat table of citations - for example, for a systematic literature review. 

The input is the location of the XML file that is obtained from EBSCOHost. 

The output is 6 columns: Record Number, Author (Year), Year, Title, Type of Publication, and hyperlinked DOI link. 

For Author (Year), there are various format depending on the number of authors (1, 2, 3+) on the reference, as follows:
1 author: Author (Year)
2 authors: Author1 & Author2 (Year)
3+ authors: Author 1 et al. (Year)

Instructions for use of the macro:

In EBSCOHost, save your citations to a folder and export to an XML file.

In cell B2 of worksheet "Parser", enter the path for the XML file

Click LoadFile button to begin parsing
