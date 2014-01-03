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

#define GSS_COLOR_CODE_LENGTH 7

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

- (void)addRow:(NSMutableArray*)entries withFormatting:(NSArray*)formatting{
  
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

- (void)addRow:(NSMutableArray*)entries withConfigurations:(NSMutableArray*)configurations withFormatting:(NSArray*)formatting{
  
  [self.array addObject:entries];
  [self.configArray addObject:configurations];
  
  //add formatting if available
  if (formatting&&formatting.count>0) {
    NSMutableArray *format = [NSMutableArray arrayWithArray:formatting];
    
    while (format.count < entries.count) {
      [format addObject:[format lastObject]];
    }
    [self.formatArray addObject:format];
  }
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

/// options for the work sheet
@property(nonatomic,retain) NSDictionary *workSheetOptions;

@end

@implementation GSSheetManager

@synthesize authorName;
@synthesize sheetArray;

/** Wrapper Method for parseStyle Defaults
 *
 *  @param styleInfo is the dictionary with all style informations
 *  @param styleID is the Style ID tag
 *  @return new style string
 */
- (NSString*)parseStyle:(NSDictionary*)styleInfo forID:(NSString*)styleID{
  return [self parseStyle:styleInfo forID:styleID forName:@""];
}

/** Helper method to parse the styleInformations
 *
 *  @param styleInfo is the dictionary with all style informations
 *  @param styleID is the Style ID tag
 *  @param name is the name flag in the style for Default
 *  @return new style string
 */
- (NSString*)parseStyle:(NSDictionary*)styleInfo forID:(NSString*)styleID forName:(NSString*)name{
  
  //invalid format will be ignored
  if ([styleInfo isKindOfClass:[NSString class]]||[styleInfo isKindOfClass:[NSNull class]]) {
    return @"";
  }
  
  //not dictionary format
  if (![styleInfo isKindOfClass:[NSDictionary class]]) {
    NSLog(@"GSSheetManager#Bad Format (%@)",NSStringFromClass([styleInfo class]));
    [NSException raise:@"BadInputException" format:@"The given input is not NSDictionary (%@)",NSStringFromClass([styleInfo class])];
  }
  
  //remove all defaults
	if ([[styleInfo allKeys] count]==0) {
		return @"";
	}
  
  //replace the all with other keys
  if ([styleInfo objectForKey:GSS_BORDER_ALL_KEY]) {
    NSMutableDictionary *dictionary =[NSMutableDictionary dictionaryWithDictionary:styleInfo];
    [dictionary setObject:[styleInfo objectForKey:GSS_BORDER_ALL_KEY] forKey:GSS_BORDER_BOTTOM_KEY];
    [dictionary setObject:[styleInfo objectForKey:GSS_BORDER_ALL_KEY] forKey:GSS_BORDER_TOP_KEY];
    [dictionary setObject:[styleInfo objectForKey:GSS_BORDER_ALL_KEY] forKey:GSS_BORDER_RIGHT_KEY];
    [dictionary setObject:[styleInfo objectForKey:GSS_BORDER_ALL_KEY] forKey:GSS_BORDER_LEFT_KEY];
    [dictionary removeObjectForKey:GSS_BORDER_ALL_KEY];
    return [self parseStyle:dictionary forID:styleID forName:name];
  }
	
	NSString *styleString = [NSString stringWithFormat:@"\t<Style ss:ID=\"%@\"%@>\n",styleID,name];
	
	//set the default font size
	NSNumber *size = [styleInfo objectForKey:@"size"];
	NSString *sizeAttribute = @"";
	if (size) {
		sizeAttribute = [NSString stringWithFormat:@" ss:Size=\"%@\"",size];
	}else{
    //try to use default style, if available
    if ((!name || [name length]==0) && self.defaultStyle) {
      NSArray *splitDefaultStyle = [self.defaultStyle componentsSeparatedByString:@" "];
      for (NSString *part in splitDefaultStyle) {
        if ([[part substringToIndex:7] isEqualToString:@"ss:Size"]) {
          sizeAttribute = [NSString stringWithFormat:@" %@",part];
          break;
        }
      }
    }
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
    if ([fontColor length]>0) {
      if ([fontColor characterAtIndex:0]!='#') {
        [NSException raise:@"BadInputException" format:@"The value for color style invalid (%@)",fontColor];
      }
      
      if ([fontColor length]>GSS_COLOR_CODE_LENGTH) {
        [NSException raise:@"BadInputException" format:@"The value for color style is too long (%@)",fontColor];
      }
    }
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
    backgroundAttribute = [NSString stringWithFormat:@"\t\t<Interior ss:Color=\"%@\" ss:Pattern=\"Solid\"/>\n",backgroundColor];
    
    if ([backgroundColor length]>0) {
      if ([backgroundColor characterAtIndex:0]!='#') {
        [NSException raise:@"BadInputException" format:@"The value for background color style invalid (%@)",backgroundColor];
      }
      
      if ([backgroundColor length]>GSS_COLOR_CODE_LENGTH) {
        [NSException raise:@"BadInputException" format:@"The value for background color style is too long (%@)",backgroundColor];
      }
    }
  }
  
  //concatenate the font string
	styleString = [NSString stringWithFormat:@"%@\t\t<ss:Font%@%@%@%@%@%@/>\n%@",
                 styleString,sizeAttribute,fontNameAttribute,fontColorAttribute,fontBoldAttribute,fontItalicAttribute,fontUnderlineAttribute,backgroundAttribute];
	
  //check wether there are alignment informations
  if (([styleInfo objectForKey:@"horizontalAlignment"]||[styleInfo objectForKey:@"verticalAlignment"])||[styleInfo objectForKey:@"wrapText"]) {
    NSString *alignmentString = @"\t\t<Alignment";
    
    //add horizontal attribute
    NSString *horizontal = [styleInfo objectForKey:@"horizontalAlignment"];
    
    if (horizontal) { // Automatic, Left, Center, Right
      alignmentString = [NSString stringWithFormat:@"%@ ss:Horizontal=\"%@\"",alignmentString,horizontal];
    }
    
    //add vertical attribute
    NSString *vertical = [styleInfo objectForKey:@"verticalAlignment"];
    
    if (vertical) { // Automatic, Top, Bottom, Center
      alignmentString = [NSString stringWithFormat:@"%@ ss:Vertical=\"%@\"",alignmentString,vertical];
    }
    
    //add wrapText attribute
    NSString *wrapText = [styleInfo objectForKey:@"wrapText"];
    
    if (wrapText) {
      alignmentString = [NSString stringWithFormat:@"%@ ss:WrapText=\"%@\"",alignmentString,wrapText];
    }
    
    alignmentString = [alignmentString stringByAppendingString:@" />\n"];
    styleString = [NSString stringWithFormat:@"%@%@",styleString,alignmentString];
  }
  
  //check wether number should be fixed
  if ([styleInfo objectForKey:GSS_NUMBER_FIXED]) {
    styleString = [NSString stringWithFormat:@"%@\t\t<NumberFormat ss:Format=\"Fixed\"/>\n",styleString];
  }
  
  //check wether border attributes exists
  if ([styleInfo objectForKey:GSS_BORDER_TOP_KEY]||
      [styleInfo objectForKey:GSS_BORDER_BOTTOM_KEY]||
      [styleInfo objectForKey:GSS_BORDER_RIGHT_KEY]||
      [styleInfo objectForKey:GSS_BORDER_LEFT_KEY]) {
    
    NSString *borderString = @"\t<Borders>\n";
    
    //border size: 0: Hairline, 1: Thin, 2: Medium, 3: Thick
    
    //top border
    if ([styleInfo objectForKey:GSS_BORDER_TOP_KEY]) {
      borderString = [NSString stringWithFormat:@"%@\t\t<Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"%@\"/>\n",borderString,[styleInfo objectForKey:GSS_BORDER_TOP_KEY]];
      
      if ([[styleInfo objectForKey:GSS_BORDER_TOP_KEY] isKindOfClass:[NSString class]]) {
        if([[styleInfo objectForKey:GSS_BORDER_TOP_KEY] length]==0){
          [NSException raise:@"BadInputException" format:@"Top Border Style is too invalid (too short)"];
        }
      }
    }
    
    //bottom border
    if ([styleInfo objectForKey:GSS_BORDER_BOTTOM_KEY]) {
      borderString = [NSString stringWithFormat:@"%@\t\t<Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"%@\"/>\n",borderString,[styleInfo objectForKey:GSS_BORDER_BOTTOM_KEY]];
      
      if ([[styleInfo objectForKey:GSS_BORDER_BOTTOM_KEY] isKindOfClass:[NSString class]]) {
        if([[styleInfo objectForKey:GSS_BORDER_BOTTOM_KEY] length]==0){
          [NSException raise:@"BadInputException" format:@"Bottom Border Style is too invalid (too short)"];
        }
      }
    }
    
    //right border
    if ([styleInfo objectForKey:GSS_BORDER_RIGHT_KEY]) {
      borderString = [NSString stringWithFormat:@"%@\t\t<Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"%@\"/>\n",borderString,[styleInfo objectForKey:GSS_BORDER_RIGHT_KEY]];
      
      if ([[styleInfo objectForKey:GSS_BORDER_RIGHT_KEY] isKindOfClass:[NSString class]]) {
        if([[styleInfo objectForKey:GSS_BORDER_RIGHT_KEY] length]==0){
          [NSException raise:@"BadInputException" format:@"Right Border Style is too invalid (too short)"];
        }
      }
    }
    
    //left border
    if ([styleInfo objectForKey:GSS_BORDER_LEFT_KEY]) {
      borderString = [NSString stringWithFormat:@"%@\t\t<Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"%@\"/>\n",borderString,[styleInfo objectForKey:GSS_BORDER_LEFT_KEY]];
      
      if ([[styleInfo objectForKey:GSS_BORDER_LEFT_KEY] isKindOfClass:[NSString class]]) {
        if([[styleInfo objectForKey:GSS_BORDER_LEFT_KEY] length]==0){
          [NSException raise:@"BadInputException" format:@"Left Border Style is too invalid (too short)"];
        }
      }
    }
    
    
    borderString = [NSString stringWithFormat:@"%@\t</Borders>\n",borderString];
    styleString = [NSString stringWithFormat:@"%@%@",styleString,borderString];
  }
  
  //add custom code
  if ([styleInfo objectForKey:GSS_CUSTOM_KEY]&&[[styleInfo objectForKey:GSS_CUSTOM_KEY] isKindOfClass:[NSString class]]) {
    styleString = [NSString stringWithFormat:@"%@\t\t%@\n",styleString,[styleInfo objectForKey:GSS_CUSTOM_KEY]];
  }
  
	//add end tag
	styleString = [NSString stringWithFormat:@"%@\t</Style>\n",styleString];
	
	return styleString;
}

- (NSString*)getStyleForDict:(NSDictionary*)style{
  
  //invalid format will be ignored
  if ([style isKindOfClass:[NSString class]]||[style isKindOfClass:[NSNull class]]) {
    return @"";
  }
  
  //not dictionary format
  if (![style isKindOfClass:[NSDictionary class]]) {
    NSLog(@"GSSheetManager#Bad Format (%@)",NSStringFromClass([style class]));
    [NSException raise:@"BadInputException" format:@"The given input is not NSDictionary (%@)",NSStringFromClass([style class])];
  }
  
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
 *  @throws BadInputException if the input is not a NSDictionary
 *  @return the Style ID as string and starting with GS
 */
- (NSString*)getStyleID:(NSDictionary*)style{
  
  //invalid format will be ignored
  if ([style isKindOfClass:[NSString class]]||[style isKindOfClass:[NSNull class]]) {
    return @"";
  }
  
  //not dictionary format
  if (![style isKindOfClass:[NSDictionary class]]) {
    NSLog(@"GSSheetManager#Bad Format (%@)",NSStringFromClass([style class]));
    [NSException raise:@"BadInputException" format:@"The given input is not NSDictionary (%@)",NSStringFromClass([style class])];
  }
  
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
	
  NSString *styleString = [self parseStyle:defaultStyle forID:@"Default" forName:@" ss:Name=\"Normal\""];
  
  //remove all defaults
	if ([styleString length]==0) {
		self.defaultStyle = nil;
	}
	
	self.defaultStyle = styleString;
}


- (void)setColumnSize:(NSArray*)list{
	self.columnSizeList = list;
}


- (void)setSheetOptions:(NSDictionary*)options{
  self.workSheetOptions = options;
}

- (NSString*)parseWorkSheetOptions:(NSDictionary*)options{
  if (options&&[options isKindOfClass:[NSDictionary class]]) {
    
    NSString *sheetOptionString = @"<WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\"><FitToPage />";
    if ([options objectForKey:GSS_WSO_ZOOM_KEY]) {
      sheetOptionString = [sheetOptionString stringByAppendingString:[NSString stringWithFormat:@"<Zoom>%@</Zoom>",[options objectForKey:GSS_WSO_ZOOM_KEY]]];
    }
    
    //optional stuff
    if ([options objectForKey:GSS_WSO_CUSTOM_KEY]) {
      sheetOptionString = [sheetOptionString stringByAppendingString:[options objectForKey:GSS_WSO_CUSTOM_KEY]];
    }
    
    return [sheetOptionString stringByAppendingString:@"</WorksheetOptions>"];
  }else{
    return @"";
  }
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
  
  //reset style counter
  _styleCounter = 0;
  NSMutableString *dataStream = [[NSMutableString alloc] init];
  
  //write meta header for spreadsheet
  [dataStream appendString:@"<?xml version=\"1.0\"?>\n<?mso-application progid=\"Excel.Sheet\"?>\n "];
  [dataStream appendString:@"<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\n xmlns:o=\"urn:schemas-microsoft-com:office:office\"\n"];
  [dataStream appendString:@"xmlns:x=\"urn:schemas-microsoft-com:office:excel\"\n xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"\n xmlns:html=\"http://www.w3.org/TR/REC-html40\">\n"];
  
  
  [dataStream appendString:@"\n<DocumentProperties xmlns=\"urn:schemas-microsoft-com:office:office\">\n"];
  
  //only set author if available
  if(self.authorName){
    [dataStream appendString:[NSString stringWithFormat:@"\t<Author>%@</Author>\n",self.authorName]];
  }
  
  //set created date
  [dataStream appendString:[NSString stringWithFormat:@"\t<Created>%@</Created>\n</DocumentProperties>\n",[NSDate date]]];
  
  //add styles for date (week day)
  [dataStream appendString:@"\n<Styles>\n"];
  
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
  
	[dataStream appendString:@"</Styles>\n\n"];
  
  //write every sheet page
  for (int i=0;i<self.sheetArray.count;i++) {
    
    GSSheetObject *sheet = [self.sheetArray objectAtIndex:i];
    
    [dataStream appendFormat:@"<Worksheet ss:Name=\"%@\">\n",sheet.sheetName];
	  
    //start the table
    [dataStream appendString:@"<Table>\n"];
	  
    //add column size if available
    if (self.columnSizeList) {
      for (int c=0; c<self.columnSizeList.count; c++) {
        [dataStream appendString:[NSString stringWithFormat:@"\t<Column ss:AutoFitWidth=\"0\" ss:Width=\"%@\"/>\n",[self.columnSizeList objectAtIndex:c]]];
      }
    }
    
    for (int j=0; j<sheet.array.count; j++) {
      //get the row
      [dataStream appendString:@"\t<Row>\n"];
      
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
          
          //only allow if there is now cellconfigurations (macros)
          if (cellConfiguration && [cellConfiguration length]>0) {
            contentData = [NSString stringWithFormat:@"%@",item];
          }else{
            contentData = item;
          }
        }else if([item isKindOfClass:[NSDate class]]){
          //item is a date
          typeString = @"DateTime";
          
          //only allow if there is now cellconfigurations (macros)
          if (!cellConfiguration) {
            contentData = [NSString stringWithFormat:@"%@",item];
          }
          
        }else if([item isKindOfClass:[NSNull class]]){
          //unknown type will treat as string
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
          if ([typeString isEqualToString:@"String"]&&[contentData length]==0) {
            //no content available
            [dataStream appendString:@"\t\t<Cell />\n"];
          }else{
            
            NSString *completeCellString = [NSString stringWithFormat:@"\t\t<Cell%@><Data ss:Type=\"%@\">%@</Data></Cell>\n",cellConfiguration,typeString,contentData];
            [dataStream appendString:completeCellString];
          }
          
        }
        
      }
      
      [dataStream appendString:@"\t</Row>\n"];
    }
    
    //end table
    [dataStream appendString:@"</Table>\n"];
    
    //if wsheet options are available add them
    if (self.workSheetOptions) {
      [dataStream appendString:[self parseWorkSheetOptions:self.workSheetOptions]];
    }
    
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
