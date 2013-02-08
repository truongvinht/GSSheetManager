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

/*! The Sheet object representates a sheet in the spreadsheet.*/
@interface GSSheetObject : NSObject

/// Name of the sheet
@property(nonatomic,strong) NSString *sheetName;

/// list with all cells in this sheet
@property(nonatomic,strong) NSMutableArray *array;

#pragma mark Methods in GSSheetObject

/** Method to init a new sheet object
 *  @param name is the title of the sheet
 */
- (id)initWithSheetName:(NSString*)name;

/** Method to add a row with entries
 *  @param entries an array with NSNumber objects as number and NSString as string
 */
- (void)addRow:(NSMutableArray*)entries;

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