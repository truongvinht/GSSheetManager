GSSheetManager
==============

An Objective-C custom Spreadsheet manager which allows to simple create tables on iOS.

#Example

The Following example creates a sheet with the name "1st page" with two rows:

Name	Hans
Points	1000

and save it into Example.xls file in the default documents folder

GSSheetManager *sheetManager = [[GSSheetManager alloc] initWithAuthor:@"Truong Vinh Tran"];

GSSheetObject *firstPage = [sheetManager addSheet:@"1st page"];

[firstPage addRow:[NSMutableArray arrayWithObjects:@"Name",@"Hans", nil]];

[firstPage addRow:[NSMutableArray arrayWithObjects:@"Points",[NSNumber numberWithDouble:1000], nil]];

[sheetManager writeSheetToFile:@"Example.xls"];