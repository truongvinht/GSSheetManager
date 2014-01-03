/*
 
 GSSheetManager.h
 
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

#ifndef GSSHEETMANAGER_H
#define GSSHEETMANAGER_H

#import <Foundation/Foundation.h>

///constants for style the cell

///font size key
#define GSS_SIZE_KEY @"size"

///font name key
#define GSS_FONTNAME_KEY @"fontName"

///font color
#define GSS_COLOR_KEY @"color"

///font is bold
#define GSS_BOLD_KEY @"bold"

///font is italic
#define GSS_ITALIC_KEY @"italic"

///font is underlined
#define GSS_UNDERLINE_KEY @"underline"

///cell background color in HEX (e.g #000000 = black)
#define GSS_BGCOLOR_KEY @"backgroundColor"

/// horizontal key
#define GSS_HORIZONTAL_KEY @"horizontalAlignment"
#define GSS_HORIZONTAL_AUTOMATIC @"Automatic"
#define GSS_HORIZONTAL_LEFT @"Left"
#define GSS_HORIZONTAL_CENTER @"Center"
#define GSS_HORIZONTAL_RIGHT @"Right"

///vertical key
#define GSS_VERTICAL_KEY @"verticalAlignment"
#define GSS_VERTICAL_AUTOMATIC @"Automatic"
#define GSS_VERTICAL_TOP @"Top"
#define GSS_VERTICAL_CENTER @"Center"
#define GSS_VERTICAL_BOTTOM @"Bottom"

///wrap the text within cell
#define GSS_WRAPTEXT_KEY @"wrapText"

///border attributes
#define GSS_BORDER_TOP_KEY @"borderTop"
#define GSS_BORDER_BOTTOM_KEY @"borderBottom"
#define GSS_BORDER_RIGHT_KEY @"borderRight"
#define GSS_BORDER_LEFT_KEY @"borderLeft"
#define GSS_BORDER_ALL_KEY @"borderAll"

///fixed number to x,xx
#define GSS_NUMBER_FIXED @"fixedNumber"

///add custom format
#define GSS_CUSTOM_KEY @"custom"

///zoom level for the work sheet window
#define GSS_WSO_ZOOM_KEY @"zoom"

///custom work sheet setting
#define GSS_WSO_CUSTOM_KEY @"custom"

/*! The Sheet object representates a sheet in the spreadsheet.*/
@interface GSSheetObject : NSObject

/// Name of the sheet
@property(nonatomic,strong) NSString *sheetName;

/// list with all cells in this sheet
@property(nonatomic,strong) NSMutableArray *array;

/// list with all configurations
@property(nonatomic,strong) NSMutableArray *configArray;

/// list with rows of formatting cells
@property(nonatomic,strong) NSMutableArray *formatArray;

#pragma mark Methods in GSSheetObject

/** Method to init a new sheet object
 *  @param name is the title of the sheet
 */
- (id)initWithSheetName:(NSString*)name;

/** Method to add a row with entries
 *  @param entries an array with NSNumber objects as number and NSString as string
 */
- (void)addRow:(NSMutableArray*)entries;

/** Method to add a row with entries and formats
 *  @param entries an array with NSNumber objects as number and NSString as string
 *  @param formatting contains cell formatting, if it contains one item, then it works for the whole row
 */
- (void)addRow:(NSMutableArray*)entries withFormatting:(NSArray*)formatting;

/** Method to add a row with entries and using configurations
 *  @param entries an array with NSNumber objects as number and NSString as string
 *  @param configurations contains links and cell configurations
 */
- (void)addRow:(NSMutableArray*)entries withConfigurations:(NSMutableArray*)configurations;

/** Method to add a row with entries and using configurations
 *  @param entries an array with NSNumber objects as number and NSString as string
 *  @param configurations contains links and cell configurations
 *  @param formatting contains cell formatting, if it contains one item, then it works for the whole row
 */
- (void)addRow:(NSMutableArray*)entries withConfigurations:(NSMutableArray*)configurations withFormatting:(NSArray*)formatting;

/** Method to replace a row with new antries
 *  @param index is the row position
 *  @param replaceArray is the new content
 */
- (void)replaceRow:(int)index withArray:(NSMutableArray*)replaceArray;

/** Method to delete a row and replace with empty entries
 *  @param index is the row index which needs to be deleted
 */
- (void)deleteRow:(int)index;

@end

/*! Class for creating Spreadsheets (using XML structure)*/
@interface GSSheetManager : NSObject

/// Array with all sheet pages
@property(nonatomic,strong) NSMutableArray *sheetArray;


/** Method to init a new object using the filename
 *  @param author is the name of the author
 *  @return the own object instance
 */
- (id)initWithAuthor:(NSString*)author;

/** Method to set the default style for the whole table
 *	@param defaultStyle is a dictionary which carries all attributes (size,fontName,color,bold,italic,underline, backgroundColor)
 */
-(void)setDefaultFontStyle:(NSDictionary*)defaultStyle;

/** Method to set the width of the column
 *	@param list is an array with NSNumbers regarding to its size
 */
- (void)setColumnSize:(NSArray*)list;

/** Method to set worksheet Options
 *  @param options is a dictionary with more options for the displaying
 */
- (void)setSheetOptions:(NSDictionary*)options;

/** Method to add a new sheetpage
 *  @param sheetName is the new name of the sheet
 *  @return GSSheetObject if the sheet was created successfully
 */
- (GSSheetObject*)addSheet:(NSString*)sheetName;

/** Method to read all sheets already in the array
 *  @return all available sheets
 */
- (NSMutableArray*)getAllSheets;

/** Method to genereate the sheet data
 *	@return the data of the sheet
 */
- (NSData*)generateSheet;

/** Method to write the given sheet to file
 *  @param fileName is the file name of the sheet
 *  @return true if the sheet was successfully written
 */
- (BOOL)writeSheetToFile:(NSString*)fileName;

/** Method to write the given sheet into the file with given path
 *  @param fileName is the file name of the sheet
 *  @param path is the location for the file
 *  @param true if the sheet was successfully written
 */
- (BOOL)writeSheetToFile:(NSString *)fileName inPath:(NSString*)path;

@end

#endif