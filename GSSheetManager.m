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
    self.formatArray = [NSMutableArray array];
  }
  return self;
}


- (void)addRow:(NSMutableArray*)entries{
  [self.array addObject:entries];
  [self.configArray addObject:[NSMutableArray array]];
  [self.formatArray addObject:[NSMutableArray array]];
}

- (void)addRow:(NSMutableArray*)entries withFormatting:(NSMutableArray*)formatting{
  
  [self.array addObject:entries];
  [self.configArray addObject:[NSMutableArray array]];
  
  //add formatting if available
  if (formatting&&formatting.count>0) {
    NSMutableArray *format = [NSMutableArray arrayWithArray:formatting];
    
    while (format.count < entries.count) {
      [format addObject:[format lastObject]];
    }
    [self.formatArray addObject:format];
  }
}

- (void)addRow:(NSMutableArray*)entries withConfigurations:(NSMutableArray*)configurations{
  [self.array addObject:entries];
  [self.configArray addObject:configurations];
  [self.formatArray addObject:[NSMutableArray array]];
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

/// column size
@property(nonatomic,retain) NSArray *columnSizeList;

/// default style
@property(nonatomic,retain) NSString *defaultStyle;

/// map for remembering available styles
@property(nonatomic,retain) NSMutableDictionary *styleMap;

/// attribute to count styles for cells
@property(nonatomic,readwrite) NSUInteger styleCounter;

@end

@implementation GSSheetManager

@synthesize authorName;
@synthesize sheetArray;

/** Helper method to parse the styleInformations
 *
 *  @param styleInfo is the dictionary with all style informations
 *  @param styleID is the Style ID tag
 */
- (NSString*)parseStyle:(NSDictionary*)styleInfo forID:(NSString*)styleID{
  
  //remove all defaults
	if ([[styleInfo allKeys] count]==0) {
		return @"";
	}
	
	NSString *styleString = [NSString stringWithFormat:@"\t<Style ss:ID=\"%@\">\n",styleID];
	
	//set the default font size
	NSNumber *size = [styleInfo objectForKey:@"size"];
	NSString *sizeAttribute = @"";
	if (size) {
		sizeAttribute = [NSString stringWithFormat:@" ss:size=\"%@\"",size];
	}
	
	//set the default fontname
	NSString *fontName = [styleInfo objectForKey:@"fontName"];
	NSString *fontNameAttribute = @"";
	if (fontName) {
		fontNameAttribute = [NSString stringWithFormat:@" ss:FontName=\"%@\" x:Family=\"Swiss\"",fontName];
	}
  
	//set the default font color
	NSString *fontColor = [styleInfo objectForKey:@"color"];
	NSString *fontColorAttribute = @"";
	if (fontColor) {
		fontColorAttribute = [NSString stringWithFormat:@" ss:Color=\"%@\"",fontColor];
	}
  
  //set the default font is Bold
	NSString *fontBold = [styleInfo objectForKey:@"bold"];
	NSString *fontBoldAttribute = @"";
	if (fontBold) {
		fontBoldAttribute = [NSString stringWithFormat:@" ss:Bold=\"1\""];
	}
  
  //set the default font is Bold
	NSString *fontItalic = [styleInfo objectForKey:@"italic"];
	NSString *fontItalicAttribute = @"";
	if (fontItalic) {
		fontItalicAttribute = [NSString stringWithFormat:@" ss:Italic=\"1\""];
	}
  
  //set the default font is underlined
	NSString *fontUnderline = [styleInfo objectForKey:@"underline"];
	NSString *fontUnderlineAttribute = @"";
	if (fontUnderline) {
		fontUnderlineAttribute = [NSString stringWithFormat:@" ss:Underline=\"1\""];
	}
  
  NSString *backgroundColor = [styleInfo objectForKey:@"backgroundColor"];
  NSString *backgroundAttribute = @"";
  if (backgroundColor) {
    backgroundAttribute = [NSString stringWithFormat:@"<Interior ss:Color=\"%@\"ss:Pattern=\"Solid\"/>\n",backgroundColor];
  }
  
  //concatenate the font string
	styleString = [NSString stringWithFormat:@"%@\t<ss:Font%@%@%@%@%@%@/>\n%@",
                 styleString,sizeAttribute,fontNameAttribute,fontColorAttribute,fontBoldAttribute,fontItalicAttribute,fontUnderlineAttribute,backgroundAttribute];
	
  //check wether border attributes exists
  if ([styleInfo objectForKey:@"borderTop"]||[styleInfo objectForKey:@"borderBottom"]||[styleInfo objectForKey:@"borderRight"]||[styleInfo objectForKey:@"borderLeft"]) {
    
    NSString *borderString = @"\t<Borders>\n";
    
    //border size: 0: Hairline, 1: Thin, 2: Medium, 3: Thick
    
    //top border
    if ([styleInfo objectForKey:@"borderTop"]) {
      borderString = [NSString stringWithFormat:@"%@\t\t<Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"%@\"/>\n",borderString,[styleInfo objectForKey:@"borderTop"]];
    }
    
    //bottom border
    if ([styleInfo objectForKey:@"borderBottom"]) {
      borderString = [NSString stringWithFormat:@"%@\t\t<Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"%@\"/>\n",borderString,[styleInfo objectForKey:@"borderBottom"]];
    }
    
    //right border
    if ([styleInfo objectForKey:@"borderRight"]) {
      borderString = [NSString stringWithFormat:@"%@\t\t<Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"%@\"/>\n",borderString,[styleInfo objectForKey:@"borderRight"]];
    }
    
    //left border
    if ([styleInfo objectForKey:@"borderLeft"]) {
      borderString = [NSString stringWithFormat:@"%@\t\t<Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"%@\"/>\n",borderString,[styleInfo objectForKey:@"borderLeft"]];
    }
    
    
    borderString = [NSString stringWithFormat:@"%@\t</Borders>\n",borderString];
    styleString = [NSString stringWithFormat:@"%@%@",styleString,borderString];
  }
  
	//add end tag
	styleString = [NSString stringWithFormat:@"%@\t</Style>\n",styleString];
	
	return styleString;
}

- (NSString*)getStyleForDict:(NSDictionary*)style{
  
  //no style
  if (!style||[[style allKeys] count]==0) {
    return nil;
  }
  
  //check wether style is already available
  for (int i=0; i<_styleCounter; i++) {
    NSArray *savedInfo = [self.styleMap objectForKey:[NSString stringWithFormat:@"GS%d",i]];
    if (savedInfo && [[savedInfo objectAtIndex:1] isEqualToDictionary:style]) {
      return nil;
    }
  }
  
  //add new style
  NSString *newStyle = [self parseStyle:style forID:[NSString stringWithFormat:@"GS%d",_styleCounter]];
  [self.styleMap setObject:@[newStyle,style] forKey:[NSString stringWithFormat:@"GS%d",_styleCounter]];
  _styleCounter++;
  return newStyle;
}

/** Method to get the style ID
 *  @param style is the dictionary with chosen style
 *  @return the Style ID as string and starting with GS
 */
- (NSString*)getStyleID:(NSDictionary*)style{
  
  //no style
  if (!style||[[style allKeys] count]==0) {
    return @"";
  }
  
  if (_styleCounter > 0) {
    //check wether style is already available
    
    for (int i=0; i<_styleCounter; i++) {
      NSArray *styleInfo = [self.styleMap objectForKey:[NSString stringWithFormat:@"GS%d",i]];
      if ([[styleInfo objectAtIndex:1] isEqualToDictionary:style]) {
        return [NSString stringWithFormat:@"GS%d",i];
      }
    }
    
  }
  //style not founded
  return @"";
}

- (id)initWithAuthor:(NSString*)author{
  self = [super init];
  if(self){
    //init the author name
    self.authorName = author;
    self.sheetArray = [NSMutableArray array];
    self.styleMap = [NSMutableDictionary dictionary];
    
    //no styles
    _styleCounter = 0;
  }
  return self;
}

-(void)setDefaultFontStyle:(NSDictionary*)defaultStyle{
	
  //remove all defaults
	if ([[defaultStyle allKeys] count]==0) {
		self.defaultStyle = nil;
	}
	
	NSString *styleString = @"\t<Style ss:ID=\"Default\" ss:Name=\"Normal\">\n";
	
	//set the default font size
	NSNumber *size = [defaultStyle objectForKey:@"size"];
	NSString *sizeAttribute = @"";
	if (size) {
		sizeAttribute = [NSString stringWithFormat:@" ss:size=\"%@\"",size];
	}
	
	//set the default fontname
	NSString *fontName = [defaultStyle objectForKey:@"fontName"];
	NSString *fontNameAttribute = @"";
	if (fontName) {
		fontNameAttribute = [NSString stringWithFormat:@" ss:FontName=\"%@\" x:Family=\"Swiss\"",fontName];
	}
  
	//set the default font color
	NSString *fontColor = [defaultStyle objectForKey:@"color"];
	NSString *fontColorAttribute = @"";
	if (fontColor) {
		fontColorAttribute = [NSString stringWithFormat:@" ss:Color=\"%@\"",fontColor];
	}
  
  //set the default font is Bold
	NSString *fontBold = [defaultStyle objectForKey:@"bold"];
	NSString *fontBoldAttribute = @"";
	if (fontBold) {
		fontBoldAttribute = [NSString stringWithFormat:@" ss:Bold=\"1\""];
	}
  
  //set the default font is Bold
	NSString *fontItalic = [defaultStyle objectForKey:@"italic"];
	NSString *fontItalicAttribute = @"";
	if (fontItalic) {
		fontItalicAttribute = [NSString stringWithFormat:@" ss:Italic=\"1\""];
	}
  
  //set the default font is underlined
	NSString *fontUnderline = [defaultStyle objectForKey:@"underline"];
	NSString *fontUnderlineAttribute = @"";
	if (fontUnderline) {
		fontUnderlineAttribute = [NSString stringWithFormat:@" ss:Underline=\"1\""];
	}
  
  NSString *backgroundColor = [defaultStyle objectForKey:@"backgroundColor"];
  NSString *backgroundAttribute = @"";
  if (backgroundColor) {
    backgroundAttribute = [NSString stringWithFormat:@"<Interior ss:Color=\"%@\"ss:Pattern=\"Solid\"/>\n",backgroundColor];
  }
  
  //concatenate the font string
	styleString = [NSString stringWithFormat:@"%@\t<ss:Font%@%@%@%@%@%@/>\n%@",
                 styleString,sizeAttribute,fontNameAttribute,fontColorAttribute,fontBoldAttribute,fontItalicAttribute,fontUnderlineAttribute,backgroundAttribute];
	
	//add end tag
	styleString = [NSString stringWithFormat:@"%@\t</Style>\n",styleString];
	
	self.defaultStyle = styleString;
}


- (void)setColumnSize:(NSArray*)list{
	self.columnSizeList = list;
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
  [dataStream appendString:@"\n<Styles>\t<Style ss:ID=\"s1\"><NumberFormat ss:Format=\"ddd\"/></Style>\n"];
  
  //add default styles if available
  if (self.defaultStyle) {
    [dataStream appendString:self.defaultStyle];
  }
  
  //add all available styles
  for (int x=0; x < self.sheetArray.count; x++) {
    
    GSSheetObject *sheet = [self.sheetArray objectAtIndex:x];
    
    for (int y=0; y < [sheet.formatArray count]; y++) {
      for (int z=0; z < [[sheet.formatArray objectAtIndex:y] count]; z++) {
        NSString *style = [self getStyleForDict:[[sheet.formatArray objectAtIndex:y] objectAtIndex:z]];
        if (style) {
          [dataStream appendString:style];
        }
      }
    }
  }
  
	[dataStream appendString:@"</Styles>\n"];
  
  //write every sheet page
  for (int i=0;i<self.sheetArray.count;i++) {
    
    GSSheetObject *sheet = [self.sheetArray objectAtIndex:i];
    
    [dataStream appendFormat:@"<Worksheet ss:Name=\"%@\">\n",sheet.sheetName];
	  
    //start the table
    [dataStream appendString:@"<Table>\n"];
	  
    //add column size if available
    if (self.columnSizeList) {
      for (int c=0; c<self.columnSizeList.count; c++) {
        [dataStream appendString:[NSString stringWithFormat:@"<Column ss:AutoFitWidth=\"0\" ss:Width=\"%@\"/>\n",[self.columnSizeList objectAtIndex:c]]];
      }
    }
    
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
        
        //format the cell
        NSMutableArray *formatRow = [sheet.formatArray objectAtIndex:j];
        if ([formatRow count]>k) {
          
          //check wether there is any style
          NSString *styleID = [self getStyleID:[formatRow objectAtIndex:k]];
          if (styleID&&[styleID length]>0) {
            cellConfiguration = [NSString stringWithFormat:@"%@ ss:StyleID=\"%@\"",cellConfiguration,styleID];
          }
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
