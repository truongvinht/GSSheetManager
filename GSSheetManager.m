/*
 
 GSSheetManager.m
 
 Copyright (c) 2013 Truong Vinh Tran
 
 Permission is hereby granted, free of charge, to any person obtaining a copy
 of this software and associated documentation files (the "Software"), to deal
 in the Software without restriction, including without limitation the rights
 to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 copies of the Software, and to permit persons to whom the Software is
 furnished to do so, subject to the following conditions:
 
 The above copyright notice and this permission notice shall be included in
 all copies or substantial portions of the Software.
 
 THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 THE SOFTWARE.
 */

#import "GSSheetManager.h"

#define GSSHEETMNGR_DEFAULT_PATH [NSSearchPathForDirectoriesInDomains(NSDocumentDirectory,NSUserDomainMask, YES) objectAtIndex:0]

@implementation GSSheetObject

@synthesize sheetName;
@synthesize array,configArray;

- (id)initWithSheetName:(NSString*)name{
  self = [super init];
  if(self){
    self.sheetName = name;
    self.array = [NSMutableArray array];
    self.configArray = [NSMutableArray array];
  }
  return self;
}


- (void)addRow:(NSMutableArray*)entries{
  [self.array addObject:entries];
  [self.configArray addObject:[NSMutableArray array]];
}

- (void)addRow:(NSMutableArray*)entries withConfigurations:(NSMutableArray*)configurations{
  [self.array addObject:entries];
  [self.configArray addObject:configurations];
}

- (void)replaceRow:(int)index withArray:(NSMutableArray*)replaceArray{
  
  //user tried to replace a new entry which is over the limit. add instead of replace
  if(index>=self.array.count){
    [self.array addObject:replaceArray];
  }else{
    [self.array replaceObjectAtIndex:index withObject:replaceArray];
  }
}

- (void)deleteRow:(int)index{
  
  //try to delete something which doesnt exist
  if(index>=self.array.count){
    //ignore it
  }else{
    [self.array replaceObjectAtIndex:index withObject:[NSMutableArray array]];
  }
}

@end


@interface GSSheetManager()

/// name of the author
@property(nonatomic,retain) NSString *authorName;

@end

@implementation GSSheetManager

@synthesize authorName;
@synthesize sheetArray;


- (id)initWithAuthor:(NSString*)author{
  self = [super init];
  if(self){
    //init the author name
    self.authorName = author;
    self.sheetArray = [NSMutableArray array];
    
  }
  return self;
}

- (GSSheetObject*)addSheet:(NSString*)sheetName{
  
  //check wether sheet name already exist
  for (GSSheetObject *sheet in self.sheetArray) {
    if([sheet.sheetName isEqualToString:sheetName])
      //sheet already exist
      return nil;
  }
  
  //create a new sheet
  GSSheetObject *sheet = [[GSSheetObject alloc] initWithSheetName:sheetName];
  [self.sheetArray addObject:sheet];
  return sheet;
}

- (NSMutableArray*)getAllSheets{
  return self.sheetArray;
}

- (NSData*)generateSheet{
  NSMutableString *dataStream = [[NSMutableString alloc] init];

  //write meta header for spreadsheet
  [dataStream appendString:@"<?xml version=\"1.0\"?>\n<?mso-application progid=\"Excel.Sheet\"?>\n "];
  [dataStream appendString:@"<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\n xmlns:o=\"urn:schemas-microsoft-com:office:office\"\n"];
  [dataStream appendString:@"xmlns:x=\"urn:schemas-microsoft-com:office:excel\"\n xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"\n xmlns:html=\"http://www.w3.org/TR/REC-html40\">\n"];


  [dataStream appendString:@"\n<DocumentProperties xmlns=\"urn:schemas-microsoft-com:office:office\">\n"];

  //only set author if available
  if(self.authorName){
    [dataStream appendString:[NSString stringWithFormat:@"<Author>%@</Author>",self.authorName]];
  }

  //set created date
  [dataStream appendString:[NSString stringWithFormat:@"<Created>%@</Created></DocumentProperties>\n",[NSDate date]]];

  //add styles for date (week day)
  [dataStream appendString:@"\n<Styles><Style ss:ID=\"s1\"><NumberFormat ss:Format=\"ddd\"/></Style></Styles>\n"];

  //write every sheet page
  for (int i=0;i<self.sheetArray.count;i++) {

    GSSheetObject *sheet = [self.sheetArray objectAtIndex:i];

    [dataStream appendFormat:@"<Worksheet ss:Name=\"%@\">\n",sheet.sheetName];

    //start the table
    [dataStream appendString:@"<Table>\n"];


    for (int j=0; j<sheet.array.count; j++) {
      //get the row
      [dataStream appendString:@"<Row>"];

      for (int k=0; k < [[sheet.array objectAtIndex:j] count]; k++) {
        //get the single cell item
        id item = [[sheet.array objectAtIndex:j] objectAtIndex:k];
        
        
        NSString *typeString = nil;
        NSString *contentData = @"";
        NSString *cellConfiguration = @"";
        
        //add cell configuration if available
        if ([sheet.array count]==[sheet.configArray count]) {
          if ([[sheet.array objectAtIndex:j] count]==[[sheet.configArray objectAtIndex:j] count]) {
            cellConfiguration = [[sheet.configArray objectAtIndex:j] objectAtIndex:k];
          }
        }
        
        //item is a string
        if([item isKindOfClass:[NSString class]]){
          typeString = @"String";
          contentData = item;
        }else if([item isKindOfClass:[NSNumber class]]){
          //item is a number
          typeString = @"Number";
          contentData = item;
        }else if([item isKindOfClass:[NSDate class]]){
          //item is a date
          typeString = @"DateTime";
          
          //only allow if there is now cellconfigurations (macros)
          if (!cellConfiguration) {
            contentData = [NSString stringWithFormat:@"%@",item];
          }
        }else if([item isKindOfClass:[NSNull class]]){
          typeString = @"String";
        }
        
        //allow short links in the cell as cell configuration
        if (typeString) {
          NSString *completeCellString = [NSString stringWithFormat:@"<Cell%@><Data ss:Type=\"%@\">%@</Data></Cell>\n",cellConfiguration,typeString,contentData];
          [dataStream appendString:completeCellString];
        }
        
      }

      [dataStream appendString:@"</Row>\n"];
    }

    //end table
    [dataStream appendString:@"</Table>\n"];

    [dataStream appendString:@"</Worksheet>\n"];
  }

  //end tag
  [dataStream appendString:@"</Workbook>"];
  return [dataStream dataUsingEncoding:NSUTF8StringEncoding];
}

- (BOOL)writeSheetToFile:(NSString*)fileName{

  //write to file
  NSString *savePath = [GSSHEETMNGR_DEFAULT_PATH stringByAppendingPathComponent:fileName];
  return [[self generateSheet] writeToFile:savePath options:NSDataWritingAtomic error:nil];
}

- (BOOL)writeSheetToFile:(NSString *)fileName inPath:(NSString*)path{
  
  //write to file
  NSString *savePath = [path stringByAppendingPathComponent:fileName];
  return [[self generateSheet] writeToFile:savePath options:NSDataWritingAtomic error:nil];
}

@end
