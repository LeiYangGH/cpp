#include <iostream>

/* run this program using the console pauser or add your own getch, system("pause") or input loop */

// MD5excelAPI.cpp : Defines the entry point for the console application.
//

//#include "stdafx.h"
// Created by Yap Chun Wei
// Version 1.0 (20 April 2006)
// Version 1.1 (22 April 2006)
	// - Fixed bugs with compound files not being able to write files more than 65535 bytes.
	// - Fixed bugs with reading and writing to Excel files containing many strings.
// Version 1.2 (30 April 2006)
	// - Added operator<< to pass BasicExcelCell to an output stream.
	// - Added Print() to BasicExcelWorksheet to print the worksheet to an output stream.
	// - Change BasicExcelCell Get functions to const functions.
	// - Rename BasicExcelWorksheet functions RenameWorkSheet() to Rename().
// Version 1.3 (10 May 2006)
	// - Fixed bugs with reading from Excel files containing Asian characters.
// Version 1.4 (13 May 2006)
	// - Fixed bugs with reading and writing to Excel files containing many strings.
// Version 1.5 (15 May 2006)
	// - Remove code for ExtSST because it was causing problems with reading and writing to Excel files containing many strings.
// Version 1.6 (16 May 2006)
	// - Optimized code for reading and writing.
// Version 1.7 (22 May 2006)
	// - Fixed code to remove some warnings.
	// - Fixed bug with BasicExcelWorksheet::Cell.
	// - Fixed bug with BasicExcel::UpdateWorksheets().
// Version 1.8 (23 May 2006)
	// - Fixed bug with reading Excel files containing many unicode strings.
	// - Fixed code to remove some warnings.
	// - Fixed variable code_ duplication in BoolErr.
	// - Minor changes to BasicExcelCell:Set functions.
// Version 1.9 (24 May 2006)
	// - Changed name_ in Style from SmallString to LargeString.
	// - Fixed bug in BasicExcelCell::GetString and BasicExcelCell::GetWString.
	// - Minor changes to functions in BasicExcel and BasicExcelWorksheet which checks for unicode.
	// - Minor change to SmallString::Read.
// Version 1.10 (30 May 2006)
	// - Fixed bug with reading Excel files containing many strings.
	// - Remove memory leaks.
// Version 1.11 (2 June 2006)
	// - Fixed bug with reading and writing Excel files containing many unicode and ANSI strings.
// Version 1.12 (6 June 2006)
	// - Fixed bug with reading and writing Excel files containing many unicode and ANSI strings.
// Version 1.13 (1 August 2006)
	// - Changed BasicExcelCell::Get() so that it will get a stored double as an integer or vice versa if necessary.
	// - Changed BasicExcelCell::Get() so that it will not cause any errors if a string is empty.
	// - Changed BasicExcelCell::SetString() and BasicExcelCell::SetWString() so that it will not save an empty string.
// Version 1.14 (6 August 2006)
	// - Fixed bug with reading Excel files that contain a null string.

#ifndef BASICEXCEL_HPP
#define BASICEXCEL_HPP

#pragma warning (disable : 4786)

#pragma warning (disable : 4305)
#pragma warning (disable : 4309)
#pragma warning (disable : 4800)
#pragma warning (disable : 4805)

#include <algorithm>
#include <cmath>
#include <functional>
#include <iostream>
#include <iomanip>
#include <fstream>
#include <map>
#include <vector>
#include <string>
#include <fstream>
#include <windows.h>
#include <string>
#include <deque>
#include <ctime>
#include <sys/stat.h>
#include <objbase.h>

using namespace std;

#define UTF16
#ifdef UTF16
	#define SIZEOFWCHAR_T 2
#else
	#define SIZEOFWCHAR_T sizeof(wchar_t)
#endif

namespace YCompoundFiles
{
class Block
// PURPOSE: In charge of handling blocks of data from a file
{
public:
	Block();

// File handling functions
	bool Create(const wchar_t* filename);
	bool Open(const wchar_t* filename, ios_base::openmode mode=ios_base::in | ios_base::out);
	bool Close();
	bool IsOpen();

// Block handling functions
	bool Read(size_t index, char* block);
	bool Write(size_t index, const char* block);
	bool Swap(size_t index1, size_t index2);
	bool Move(size_t from, size_t to);
	bool Insert(size_t index, const char* block);
	bool Erase(size_t index);
	bool Erase(vector<size_t>& indices);

// Misc functions
	size_t GetBlockSize() const {return blockSize_;}
	void SetBlockSize(size_t size) 
	{
		blockSize_ = size;
		indexEnd_ = fileSize_/blockSize_ + (fileSize_ % blockSize_ ? 1 : 0);
	}
	
protected:
	vector<char> filename_;
	ios_base::openmode mode_;
	fstream file_;
	size_t blockSize_;
	size_t indexEnd_;
	size_t fileSize_;
};

struct LittleEndian
{
	#define READWRITE(Type) \
	static void Read(const char* buffer, Type& retVal, int pos=0, int bytes=0)	\
	{	\
		retVal = Type(0);	\
		if (bytes == 0) bytes = sizeof(Type);	\
		for (size_t i=0; i<bytes; ++i)	\
		{	\
			retVal |= ((Type)((unsigned char)buffer[pos+i])) << 8*i;	\
		}	\
	}	\
	static void ReadString(const char* buffer, Type* str, int pos=0, int bytes=0)	\
	{	\
		for (size_t i=0; i<bytes; ++i) Read(buffer, str[i], pos+i*sizeof(Type));	\
	}	\
	static void Write(char* buffer, Type val, int pos=0, int bytes=0)	\
	{	\
		if (bytes == 0) bytes = sizeof(Type);	\
		for (size_t i=0; i<bytes; ++i)	\
		{	\
			buffer[pos+i] = (unsigned char)val;	\
			val >>= 8;	\
		}	\
	}	\
	static void WriteString(char* buffer, Type* str, int pos=0, int bytes=0)	\
	{	\
		for (size_t i=0; i<bytes; ++i) Write(buffer, str[i], pos+i*sizeof(Type));	\
	}	\
	static void Read(const vector<char>& buffer, Type& retVal, int pos=0, int bytes=0)	\
	{	\
		retVal = Type(0);	\
		if (bytes == 0) bytes = sizeof(Type);	\
		for (size_t i=0; i<bytes; ++i)	\
		{	\
			retVal |= ((Type)((unsigned char)buffer[pos+i])) << 8*i;	\
		}	\
	}	\
	static void ReadString(const vector<char>& buffer, Type* str, int pos=0, int bytes=0)	\
	{	\
		for (size_t i=0; i<bytes; ++i) Read(buffer, str[i], pos+i*sizeof(Type));	\
	}	\
	static void Write(vector<char>& buffer, Type val, int pos=0, int bytes=0)	\
	{	\
		if (bytes == 0) bytes = sizeof(Type);	\
		for (size_t i=0; i<bytes; ++i)	\
		{	\
			buffer[pos+i] = (unsigned char)val;	\
			val >>= 8;	\
		}	\
	}	\
	static void WriteString(vector<char>& buffer, Type* str, int pos=0, int bytes=0)	\
	{	\
		for (size_t i=0; i<bytes; ++i) Write(buffer, str[i], pos+i*sizeof(Type));	\
	}	\
	
	READWRITE(char)
	READWRITE(unsigned char)
	READWRITE(short)
	READWRITE(int)
	READWRITE(unsigned int)
	READWRITE(long)
	READWRITE(unsigned long)
	READWRITE(__int64)
	READWRITE(unsigned __int64)

	#undef READWRITE


	static void Read(const char* buffer, wchar_t& retVal, int pos=0, int bytes=0)
	{
		retVal = wchar_t(0);
		if (bytes == 0) bytes = SIZEOFWCHAR_T;
		for (int i=0; i<bytes; ++i)
		{
			retVal |= ((wchar_t)((unsigned char)buffer[pos+i])) << 8*i;
		}
	}

	static void ReadString(const char* buffer, wchar_t* str, int pos=0, int bytes=0)
	{
		for (int i=0; i<bytes; ++i) Read(buffer, str[i], pos+i*SIZEOFWCHAR_T);
	}

	static void Write(char* buffer, wchar_t val, int pos=0, int bytes=0)
	{
		if (bytes == 0) bytes = SIZEOFWCHAR_T;
		for (int i=0; i<bytes; ++i)
		{
			buffer[pos+i] = (unsigned char)val;
			val >>= 8;
		}
	}

	static void WriteString(char* buffer, wchar_t* str, int pos=0, int bytes=0)
	{
		for (int i=0; i<bytes; ++i) Write(buffer, str[i], pos+i*SIZEOFWCHAR_T);
	}

	static void Read(const vector<char>& buffer, wchar_t& retVal, int pos=0, int bytes=0)
	{
		retVal = wchar_t(0);
		if (bytes == 0) bytes = SIZEOFWCHAR_T;
		for (int i=0; i<bytes; ++i)
		{
			retVal |= ((wchar_t)((unsigned char)buffer[pos+i])) << 8*i;
		}
	}

	static void ReadString(const vector<char>& buffer, wchar_t* str, int pos=0, int bytes=0)
	{
		for (int i=0; i<bytes; ++i) Read(buffer, str[i], pos+i*SIZEOFWCHAR_T);
	}

	static void Write(vector<char>& buffer, wchar_t val, int pos=0, int bytes=0)
	{
		if (bytes == 0) bytes = SIZEOFWCHAR_T;
		for (int i=0; i<bytes; ++i)
		{
			buffer[pos+i] = (unsigned char)val;
			val >>= 8;
		}
	}

	static void WriteString(vector<char>& buffer, wchar_t* str, int pos=0, int bytes=0)
	{
		for (int i=0; i<bytes; ++i) Write(buffer, str[i], pos+i*SIZEOFWCHAR_T);
	}
};

class CompoundFile
{
public:
	enum {DUPLICATE_PROPERTY=-6,
		  NAME_TOO_LONG=-5, FILE_NOT_FOUND=-4, 
		  DIRECTORY_NOT_EMPTY=-3, DIRECTORY_NOT_FOUND=-2, 
		  INVALID_PATH=-1, 
		  SUCCESS=1};

	CompoundFile();
	~CompoundFile();

// User accessible functions
public:
	// Compound File functions
	bool Create(const wchar_t* filename);
	bool Open(const wchar_t* filename, ios_base::openmode mode=ios_base::in | ios_base::out);
	bool Close();
	bool IsOpen();

	// Directory functions
	int ChangeDirectory(const wchar_t* path);
	int MakeDirectory(const wchar_t* path);
	int PresentWorkingDirectory(wchar_t* path);
	int PresentWorkingDirectory(vector<wchar_t>& path);
	int RemoveDirectory(const wchar_t* path);
	int DelTree(const wchar_t* path);
	int DirectoryList(vector<vector<wchar_t> >& list, const wchar_t* path=0);

	// File functions
	int MakeFile(const wchar_t* path);
	int RemoveFile(const wchar_t* path);
	int FileSize(const wchar_t* path, size_t& size);
	int ReadFile(const wchar_t* path, char* data);
	int ReadFile(const wchar_t* path, vector<char>&data);
	int WriteFile(const wchar_t* path, const char* data, size_t size);
	int WriteFile(const wchar_t* path, const vector<char>&data, size_t size);


	// ANSI char functions
	bool Create(const char* filename);
	bool Open(const char* filename, ios_base::openmode mode=ios_base::in | ios_base::out);
	int ChangeDirectory(const char* path);
	int MakeDirectory(const char* path);
	int PresentWorkingDirectory(char* path);
	int PresentWorkingDirectory(vector<char>& path);
	int RemoveDirectory(const char* path);
	int DelTree(const char* path);
	int MakeFile(const char* path);
	int RemoveFile(const char* path);
	int FileSize(const char* path, size_t& size);
	int ReadFile(const char* path, char* data);
	int ReadFile(const char* path, vector<char>& data);
	int WriteFile(const char* path, char* data, size_t size);
	int WriteFile(const char* path, vector<char>& data, size_t size);

// Protected functions and data members
protected:
	// General functions and data members
	void IncreaseLocationReferences(vector<size_t> indices);
	void DecreaseLocationReferences(vector<size_t> indices);
	void SplitPath(const wchar_t* path, wchar_t*& parentpath, wchar_t*& propertyname);
	vector<char> block_;
	Block file_;

	// Header related functions and data members
	bool LoadHeader();
	void SaveHeader();
	class Header
	{
	public:
		Header();
		void Write(char* block);
		void Read(char* block);

		__int64 fileType_;		// Magic number identifying this as a compound file system (0x0000)
		int uk1_;					// Unknown constant (0x0008)
		int uk2_;					// Unknown constant (0x000C)
		int uk3_;					// Unknown constant (0x0010)
		int uk4_;					// Unknown constant (0x0014)
		short uk5_;					// Unknown constant (revision?) (0x0018)
		short uk6_;					// Unknown constant	(version?) (0x001A)
		short uk7_;					// Unknown constant (0x001C)
		short log2BigBlockSize_;	// Log, base 2, of the big block size (0x001E)
		int log2SmallBlockSize_;	// Log, base 2, of the small block size (0x0020)
		int uk8_;					// Unknown constant (0x0024)
		int uk9_;					// Unknown constant (0x0028)
		int BATCount_;				// Number of elements in the BAT array (0x002C)
		int propertiesStart_;		// Block index of the first block of the property table (0x0030)
		int uk10_;					// Unknown constant (0x0034)
		int uk11_;					// Unknown constant (0x0038)
		int SBATStart_;				// Block index of first big block containing the small block allocation table (SBAT) (0x003C)
		int SBATCount_;				// Number of big blocks holding the SBAT (0x0040)
		int XBATStart_;				// Block index of the first block in the Extended Block Allocation Table (XBAT) (0x0044)
		int XBATCount_;				// Number of elements in the Extended Block Allocation Table (to be added to the BAT) (0x0048)
		int BATArray_[109];			// Array of block indices constituting the Block Allocation Table (BAT) (0x004C, 0x0050, 0x0054 ... 0x01FC)

		size_t bigBlockSize_;
		size_t smallBlockSize_;

	private:
		void Initialize();
	};	
	Header header_;

	// BAT related functions and data members
	void LoadBAT();
	void SaveBAT();
	size_t DataSize(size_t startIndex, bool isBig);
	size_t ReadData(size_t startIndex, char* data, bool isBig);
	size_t WriteData(const char* data, size_t size, int startIndex, bool isBig);
	void GetBlockIndices(size_t startIndex, vector<size_t>& indices, bool isBig);
	size_t GetFreeBlockIndex(bool isBig);
	void ExpandBATArray(bool isBig);
	void LinkBlocks(size_t from, size_t to, bool isBig);
	void FreeBlocks(vector<size_t>& indices, bool isBig);
	vector<int> blocksIndices_;
	vector<int> sblocksIndices_;	

	// Properties related functions and data members
	class Property
	{
	public:
		Property();		
		void Write(char* block);
		void Read(char* block);
		friend bool operator==(const Property& lhs, const Property& rhs) 
		{
			return (!wcscmp(lhs.name_, rhs.name_));
		}
		friend bool operator< (const Property& lhs, const Property& rhs)
		{
			size_t maxLen1 = wcslen(lhs.name_);
			size_t maxLen2 = wcslen(rhs.name_);
			if (maxLen1 < maxLen2) return true;
			else if (maxLen1 > maxLen2) return false;
			else
			{
				int result = wcscmp(lhs.name_, rhs.name_);
				if (result <= 0) return true;
				else return false;
			}
		}
		friend bool operator!=(const Property& lhs, const Property& rhs) {return !(lhs == rhs);}
		friend bool operator> (const Property& lhs, const Property& rhs) {return (rhs < lhs);}
		friend bool operator<=(const Property& lhs, const Property& rhs) {return !(rhs < lhs);}
		friend bool operator>=(const Property& lhs, const Property& rhs) {return !(lhs < rhs);}

		wchar_t name_[32];				// A unicode null-terminated uncompressed 16bit string (lblocke the high bytes) containing the name of the property. (0x00, 0x02, 0x04, ... 0x3E)
		short nameSize_;				// Number of characters in the NAME field (0x40)
		unsigned char propertyType_;	// Property type (directory, file, or root) Byte 1 (directory), 2 (file), or 5 (root entry) (0x42)
		unsigned char nodeColor_;		// Node color (0x43)
		int previousProp_;				// Previous property index (0x44)
		int nextProp_;					// Next property index (0x48)
		int childProp_;					// First child property index (0x4c)
		int uk1_;
		int uk2_;
		int uk3_;
		int uk4_;
		int uk5_;
		int seconds1_;					// Seconds component of the created timestamp? (0x64)
		int days1_;						// Days component of the created timestamp? (0x68)
		int seconds2_;					// Seconds component of the modified timestamp? (0x6C)
		int days2_;						// Days component of the modified timestamp? (0x70)
		int startBlock_;				// Starting block of the file, used as the first block in the file and the pointer to the next block from the BAT (0x74)
		int size_;						// Actual size of the file this property points to. (used to truncate the blocks to the real size). (0x78)
	};
	class PropertyTree
	{
	public:
		PropertyTree();
		~PropertyTree();
		PropertyTree* parent_;
		Property* self_;
		size_t index_;
		vector<PropertyTree*> children_;
	};
	void LoadProperties();
	void SaveProperties();
	int MakeProperty(const wchar_t* path, Property* property);
	PropertyTree* FindProperty(size_t index);
	PropertyTree* FindProperty(const wchar_t* path);
	PropertyTree* FindProperty(PropertyTree* parentTree, wchar_t* name);
	void InsertPropertyTree(PropertyTree* parentTree, Property* property, size_t index);
	void DeletePropertyTree(PropertyTree* tree);
	void UpdateChildrenIndices(PropertyTree* parentTree);
	void IncreasePropertyReferences(PropertyTree* parentTree, size_t index);
	void DecreasePropertyReferences(PropertyTree* parentTree, size_t index);
	PropertyTree* propertyTrees_;
	PropertyTree* currentDirectory_;
	vector<Property*> properties_;
	vector<PropertyTree*> previousDirectories_;
};
} // YCompoundFiles namespace end

namespace YExcel
{
using namespace YCompoundFiles;

struct CODE
{
	enum {	FORMULA=0x0006,		//Token array and the result of a formula cell.
			YEOF=0x000A,			//End of a record block with leading BOF record.
			CALCCOUNT=0x000C,	//Maximum number of times the forumlas should be iteratively calculated
			CALCMODE=0x000D,	//Calculate formulas manually, automatically, or automatically except for multiple table operations
			PRECISION=0x000E,	//Whether formulas use the real cell values for calculation or the values displayed on the screen.
			REFMODE=0x000F,		//Method used to show cell addresses in formulas.
			DELTA=0x0010,		//Maximum change of the result to exit an iteration.
			ITERATION=0x0011,	//Whether iterations are allowed while calculating recursive formulas.
			PROTECT=0x0012,		//Whether worksheet or a workbook is protected against modification.
			PASSWORD=0x0013,	//16-bit hash value, calculated from the worksheet or workbook protection password.
			HEADER=0x0014,		//Page header string for the current worksheet.
			FOOTER=0x0015,		//Page footer string for the current worksheet.
			EXTERNSHEET=0x0017,	//List with indexes to SUPBOOK records
			NAME=0x0018,		//Name and token array of an internal defined name.
			WINDOWPROTECT=0x0019,	//Whether the window configuration of the document is protected.
			SELECTION=0x001D,	//Addresses of all selected cell ranges and position of the active cell for a pane in the current sheet.
			DATEMODE=0x0022,	//Base date for displaying date values.
			EXTERNNAME=0x0023,	//Name of an external defined name, name of an add-in function, a DDE item or an OLE object storage identifier.
			LEFTMARGIN=0x0026,	//Left page margin of the current worksheet.
			RIGHTMARGIN=0x0027,	//Right page margin of the current worksheet.
			TOPMARGIN=0x0028,	//Top page margin of the current worksheet.
			BOTTOMMARGIN=0x0029,	//Bottom page margin of current worksheet
			PRINTHEADERS=0x002A,	//Whether row and column headers (the areas with row numbers and column letters) will be printed.
			PRINTGRIDLINES=0x002B,	//Whether sheet grid lines will be printed.
			FILEPASS=0x002F,	//Information about the read/write password of the file.
			FONT=0x0031,		//Information about a used font, including character formatting.
			TABLE=0x0036,		//Information about a multiple operation table in the sheet.
			CONTINUE=0x003C,	//Continue from previous record
			WINDOW1=0x003D,		//General settings for the workbook global settings.
			BACKUP=0x0040,		//Make backup of file while saving?
			PANE=0x0041,		//Position of window panes.
			CODEPAGE=0x0042,	//Text encoding used to encode byte strings
			DCONREF=0x0051,
			DEFCOLWIDTH=0x0055,	//Default column width for columns that do not have a specific width set
			XCT=0x0059,			//Number of immediately following CRN records.
			CRN=0x005A,			//Contents of an external cell or cell range.
			FILESHARING=0x005B,	//Information about write protection, for instance the write protection password.
			WRITEACCESS=0x005C,	//Name of the user that has saved the file.
			UNCALCED=0x005E,	//Formulas have not been recalculated before the document was saved.
			SAVERECALC=0x005F,	//"Recalculate before save" option
			OBJECTPROTECT=0x0063,	//Whether objects of the current sheet are protected.
			COLINFO=0x007D,		//Width for a given range of columns
			GUTS=0x0080,		//Layout of outline symbols.
			WSBOOL=0x0081,		//16-bit value with boolean options for the current sheet.
			GRIDSET=0x0082,		//Whether option to print sheet grid lines has ever been changed.
			HCENTER=0x0083,		//Sheet is centred horizontally when printed.
			VCENTER=0x0084,		//Whether sheet is centred vertically when printed.
			BOUNDSHEET=0x0085,	//Sheet inside of the workbook
			WRITEPROT=0x0086,	//Whether file is write protected.
			COUNTRY=0x008C,		//User interface language of the Excel version that has saved the file, system regional settings at the time the file was saved.
			HIDEOBJ=0x008D,		//Whether and how to show objects in the workbook.
			SORT=0x0090,		//Last settings from the "Sort" dialogue for each sheet.
			PALETTE=0x0092,		//Definition of all user-defined colours available for cell and object formatting.
			SETUP=0x00A1,		//Page format settings of the current sheet.
			SHRFMLA=0x00BC,		//Token array of a shared formula.
			MULRK=0x00BD,		//Cell range containing RK value cells. All cells are located in the same row.
			MULBLANK=0x00BE,	//Cell range of empty cells. All cells are located in the same row.
			DBCELL=0x00D7,		//Relative offsets to calculate stream position of the first cell record for each row.
			BOOKBOOL=0x00DA,	//Save values linked from external workbooks records and XCT records?
			SCENPROTECT=0x00DD,	//Whether scenarios of the current sheet are protected.
			XF=0x00E0,			//Formatting information for cells, rows, columns or styles.
			MERGEDCELLS=0x00E5,	//All merged cell ranges of the current sheet.
			SST=0x00FC,			//List of all strings used anywhere in the workbook.
			LABELSST=0x00FD,	//Cell that contains a string.
			EXTSST=0x00FF,		//Create a hash table with stream offsets to the SST record to optimise string search operations.
			LABELRANGES=0x015F,	//Addresses of all row and column label ranges in the current sheet.
			USESELFS=0x0160,	//Whether formulas in the workbook can use "natural language formulas".
			DSF=0x0161,			//Whether file contains an addition BIFF5/BIFF7 workbook stream.
			SUPBOOK=0x01AE,		//URL of an external document and a list of sheet names inside this document.
			CONDFMT=0x01B0,		//List of cell range addresses for all cells with equal conditional formatting.
			CF=0x01B1,			//Condition and the formatting attributes applied to the cells specified in the CONDFMT record, if the condition is met
			DVAL=0x01B2,		//List header of the data validity table in the current sheet.
			HLINK=0x01B8,		//One cell address or a cell range where all cells contain the same hyperlink.
			DV=0x01BE,			//Data validity settings and a list of cell ranges which contain these settings.
			DIMENSIONS=0x0200,	//Range address of the used area in the current sheet.
			BLANK=0x0201,		//Empty cell, contains cell address and formatting information
			NUMBER=0x0203,		//Cell that contains a floating-point value.
			BOOLERR=0x0205,		//Error value cell
			STRING=0x0207,		//Result of a string formula.
			ROW=0x0208,			//Properties of a single row in the sheet.
			INDEX=0x020B,		//Range of used rows and stream positions of several records of the current sheet.
			ARRAY=0x0221,		//Token array of an array formula
			WINDOW2=0x023E,		//Additional settings for the window of a specific worksheet.
			RK=0x027E,			//Cell that contains an RK value (encoded integer or floating point value).
			STYLE=0x0293,		//Name of a user-defined cell style or specific options for a built-in cell style.
			FORMAT=0x041E,		//Number format.
			SHRFMLA1=0x04BC,	//Token array of a shared formula (added).
			QUICKTIP=0x0800,	//Cell range and text for a tool tip.
			BOF=0x0809,			//Beginning of file
			SHEETLAYOUT=0x0862,	//Colour of the tab below the sheet containing the sheet name.
			SHEETPROTECTION=0x0867,	//Additional options for sheet protection.
			RANGEPROTECTION=0x0868	//Information about special protected ranges in a protected sheet.
		};
};

class Record
{
public:
	Record();
	virtual ~Record();
	virtual size_t Read(const char* data);
	virtual size_t Write(char* data);	
	virtual size_t DataSize();
	virtual size_t RecordSize();
	short code_;
	vector<char> data_;
	size_t dataSize_;
	size_t recordSize_;
	vector<size_t> continueIndices_;
};

struct BOF : public Record
{
	BOF();
	virtual size_t Read(const char* data);
	virtual size_t Write(char* data);	
	short version_;
	short type_;
	short buildIdentifier_;
	short buildYear_;
	int fileHistoryFlags_;
	int lowestExcelVersion_;
};

struct YEOF : public Record
{
	YEOF();
};

struct SmallString
{
	SmallString();
	~SmallString();
	SmallString(const SmallString& s);
	SmallString& operator=(const SmallString& s);
	const SmallString& operator=(const char* str);
	const SmallString& operator=(const wchar_t* str);
	void Reset();
	size_t Read(const char* data);
	size_t Write(char* data);	
	size_t DataSize();
	size_t RecordSize();
	size_t StringSize();
	wchar_t* wname_;
	char* name_;
	char unicode_;
};

struct LargeString
{
	LargeString();
	~LargeString();
	LargeString(const LargeString& s);
	LargeString& operator=(const LargeString& s);
	const LargeString& operator=(const char* str);
	const LargeString& operator=(const wchar_t* str);
	void Reset();
	size_t Read(const char* data);
	size_t ContinueRead(const char* data, size_t size);
	size_t Write(char* data);	
	size_t DataSize();
	size_t RecordSize();
	size_t StringSize();
	vector<wchar_t> wname_;
	vector<char> name_;
	char unicode_;
	short richtext_;
	int phonetic_;
};

class Workbook
{
public:
	Workbook();

public:
	struct FileProtection;
	struct CodePage;
	struct DSF;
	struct TabID;
	struct FnGroupCount;
	struct WorkbookProtection;
	struct Window1 : public Record
	{
		Window1();
		virtual size_t Read(const char* data);
		virtual size_t Write(char* data);	
		short horizontalPos_;
		short verticalPos_;
		short width_;
		short height_;
		short options_;
		short activeWorksheetIndex_;
		short firstVisibleTabIndex_;
		short selectedWorksheetNo_;
		short worksheetTabBarWidth_;
	};
	struct Backup;
	struct HideObj;
	struct DateMode;
	struct Precision;
	struct RefreshAll;
	struct BookBool;
	struct Font : public Record
	{
		Font();
		virtual size_t Read(const char* data);
		virtual size_t Write(char* data);	
		virtual size_t DataSize();
		virtual size_t RecordSize();
		short height_;
		short options_;
		short colourIndex_;
		short weight_;
		short escapementType_;
		char underlineType_;
		char family_;
		char characterSet_;
		char unused_;
		SmallString name_;
	};
	struct Format;
	struct XF : public Record
	{
		XF();
		virtual size_t Read(const char* data);
		virtual size_t Write(char* data);	
		short fontRecordIndex_;
		short formatRecordIndex_;
		short protectionType_;
		char alignment_;
		char rotation_;
		char textProperties_;
		char usedAttributes_;
		int borderLines_;
		int colour1_;
		short colour2_;
	};
	struct Style : public Record
	{
		Style();
		virtual size_t Read(const char* data);
		virtual size_t Write(char* data);	
		virtual size_t DataSize();
		virtual size_t RecordSize();
		short XFRecordIndex_;
		char identifier_;
		char level_;
		LargeString name_;		
	};
	struct Palette;
	struct UseSelfs;
	struct BoundSheet : public Record
	{
		BoundSheet();
		virtual size_t Read(const char* data);
		virtual size_t Write(char* data);	
		virtual size_t DataSize();
		virtual size_t RecordSize();
		int BOFpos_;
		char visibility_;
		char type_;
		SmallString name_;		
	};
	struct Country;
	struct LinkTable;
	struct SharedStringTable : public Record
	{
		SharedStringTable();
		virtual size_t Read(const char* data);
		virtual size_t Write(char* data);	
		virtual size_t DataSize();
		virtual size_t RecordSize();
		int stringsTotal_;
		int uniqueStringsTotal_;
		vector<LargeString> strings_;	
	};
	struct ExtSST : public Record
	{
		ExtSST();
		virtual size_t Read(const char* data);
		virtual size_t Write(char* data);
		virtual size_t DataSize();
		virtual size_t RecordSize();
		short stringsTotal_;
		vector<int> streamPos_;
		vector<short> firstStringPos_;
		vector<short> unused_;
	};
	size_t Read(const char* data);
	size_t Write(char* data);
	size_t DataSize();
	size_t RecordSize();

	BOF bof_;
	Window1 window1_;
	vector<Font> fonts_;
	vector<XF> XFs_;
	vector<Style> styles_;
	vector<BoundSheet> boundSheets_;
	SharedStringTable sst_;
	ExtSST extSST_;
	YEOF eof_;
};

class Worksheet
{
public:
	Worksheet();

public:
	struct Uncalced;
	struct Index : public Record
	{
		Index();
		virtual size_t Read(const char* data);
		virtual size_t Write(char* data);	
		virtual size_t DataSize();
		virtual size_t RecordSize();
		int unused1_;
		size_t firstUsedRowIndex_;
		size_t firstUnusedRowIndex_;
		int unused2_;
		vector<size_t> DBCellPos_;
		
	};
	struct CalculationSettings
	{
		struct CalcCount;
		struct CalcMode;
		struct RefMode;
		struct Delta;
		struct Iteration;
		struct SafeRecalc;		
	};
	struct PrintHeaders;
	struct PrintGridlines;
	struct Gridset;
	struct Guts;
	struct DefaultRowHeight;
	struct WSBool;
	struct PageSettings
	{
		struct Header;
		struct Footer;
		struct HCenter;
		struct VCenter;
		struct LeftMargin;
		struct RightMargin;
		struct TopMargin;
		struct BottomMargin;
		struct PLS;
		struct Setup;
	};
	struct WorksheetProtection;
	struct DefColWidth;
	struct ColInfo;
	struct Sort;
	struct Dimensions : public Record
	{
		Dimensions();
		virtual size_t Read(const char* data);
		virtual size_t Write(char* data);	
		size_t firstUsedRowIndex_;
		size_t lastUsedRowIndexPlusOne_;
		short firstUsedColIndex_;
		short lastUsedColIndexPlusOne_;
		short unused_;		
	};
	struct CellTable
	{
		struct RowBlock
		{
			struct Row : public Record
			{
				Row();
				virtual size_t Read(const char* data);
				virtual size_t Write(char* data);	
				short rowIndex_;
				short firstCellColIndex_;
				short lastCellColIndexPlusOne_;
				short height_;
				short unused1_;
				short unused2_;
				int options_;
			};
			struct CellBlock
			{
				struct Blank : public Record
				{
					Blank();
					virtual size_t Read(const char* data);
					virtual size_t Write(char* data);	
					short rowIndex_;
					short colIndex_;
					short XFRecordIndex_;					
				};
				struct BoolErr : public Record
				{
					BoolErr();
					virtual size_t Read(const char* data);
					virtual size_t Write(char* data);	
					short rowIndex_;
					short colIndex_;
					short XFRecordIndex_;
					char value_;
					char error_;	
				};
				struct LabelSST : public Record
				{
					LabelSST();
					virtual size_t Read(const char* data);
					virtual size_t Write(char* data);	
					short rowIndex_;
					short colIndex_;
					short XFRecordIndex_;
					size_t SSTRecordIndex_;	
				};
				struct MulBlank : public Record
				{
					MulBlank();
					virtual size_t Read(const char* data);
					virtual size_t Write(char* data);	
					virtual size_t DataSize();
					virtual size_t RecordSize();
					short rowIndex_;
					short firstColIndex_;
					vector<short> XFRecordIndices_;
					short lastColIndex_;
				};
				struct MulRK : public Record
				{
					MulRK();
					virtual size_t Read(const char* data);
					virtual size_t Write(char* data);
					virtual size_t DataSize();
					virtual size_t RecordSize();
					struct XFRK
					{
						XFRK();
						void Read(const char* data);
						void Write(char* data);
						short XFRecordIndex_;
						int RKValue_;
					};
					short rowIndex_;
					short firstColIndex_;
					vector<XFRK> XFRK_;
					short lastColIndex_;
				};
				struct Number : public Record
				{
					Number();
					virtual size_t Read(const char* data);
					virtual size_t Write(char* data);	
					short rowIndex_;
					short colIndex_;
					short XFRecordIndex_;
					double value_;

				private:
					union 
					{
						__int64 intvalue_;
						double doublevalue_;
					} intdouble_;
				};
				struct RK : public Record
				{
					RK();
					virtual size_t Read(const char* data);
					virtual size_t Write(char* data);	
					short rowIndex_;
					short colIndex_;
					short XFRecordIndex_;
					int value_;	
				};

				struct Formula : public Record
				{
					struct Array : public Record
					{
						Array();
						virtual size_t Read(const char* data);
						virtual size_t Write(char* data);	
						virtual size_t DataSize();
						virtual size_t RecordSize();
						short firstRowIndex_;
						short lastRowIndex_;
						char firstColIndex_;
						char lastColIndex_;
						short options_;
						int unused_;
						vector<char> formula_;
					};
					struct ShrFmla : public Record
					{
						ShrFmla();
						virtual size_t Read(const char* data);
						virtual size_t Write(char* data);	
						virtual size_t DataSize();
						virtual size_t RecordSize();
						short firstRowIndex_;
						short lastRowIndex_;
						char firstColIndex_;
						char lastColIndex_;
						short unused_;
						vector<char> formula_;
					};
					struct ShrFmla1 : public Record
					{
						ShrFmla1();
						virtual size_t Read(const char* data);
						virtual size_t Write(char* data);	
						virtual size_t DataSize();
						virtual size_t RecordSize();
						short firstRowIndex_;
						short lastRowIndex_;
						char firstColIndex_;
						char lastColIndex_;
						short unused_;
						vector<char> formula_;
					};
					struct Table : public Record
					{
						Table();
						virtual size_t Read(const char* data);
						virtual size_t Write(char* data);	
						short firstRowIndex_;
						short lastRowIndex_;
						char firstColIndex_;
						char lastColIndex_;
						short options_;
						short inputCellRowIndex_;
						short inputCellColIndex_;
						short inputCellColumnInputRowIndex_;
						short inputCellColumnInputColIndex_;
					};
					struct String : public Record
					{
						String();
						virtual size_t Read(const char* data);
						virtual size_t Write(char* data);	
						virtual size_t DataSize();
						virtual size_t RecordSize();
						vector<wchar_t> string_;
					};

					Formula();
					virtual size_t Read(const char* data);
					virtual size_t Write(char* data);	
					virtual size_t DataSize();
					virtual size_t RecordSize();
					short rowIndex_;
					short colIndex_;
					short XFRecordIndex_;
					char result_[8];
					short options_;
					int unused_;
					vector<char> RPNtoken_;
					short type_;

					Array array_;
					ShrFmla shrfmla_;
					ShrFmla1 shrfmla1_;
					Table table_;
					String string_;
				};

				CellBlock();
				~CellBlock();
				void Reset();
				size_t Read(const char* data);
				size_t Write(char* data);
				size_t DataSize();
				size_t RecordSize();
				short RowIndex();
				short ColIndex();
				short LastColIndex();
				short type_;
				bool normalType_;
				
				Blank blank_;
				BoolErr boolerr_;
				LabelSST labelsst_;
				MulBlank mulblank_;
				MulRK mulrk_;
				Number number_;
				RK rk_;
				Formula formula_;
			};
			struct DBCell : public Record
			{
				DBCell();
				virtual size_t Read(const char* data);
				virtual size_t Write(char* data);	
				virtual size_t DataSize();
				virtual size_t RecordSize();
				int firstRowOffset_;
				vector<short> offsets_;
			};			
			
			size_t Read(const char* data);
			size_t Write(char* data);
			size_t DataSize();
			size_t RecordSize();
		
			vector<Row> rows_;
			vector<CellBlock> cellBlocks_;
			DBCell dbcell_;
		};	
		size_t Read(const char* data);
		size_t Write(char* data);
		size_t DataSize();
		size_t RecordSize();
	
		vector<RowBlock> rowBlocks_;		
	};
	struct Window2 : public Record
	{
		Window2();
		virtual size_t Read(const char* data);
		virtual size_t Write(char* data);	
		short options_;
		short firstVisibleRowIndex_;
		short firstVisibleColIndex_;
		short gridLineColourIndex_;
		short unused1_;
		short magnificationFactorPageBreakPreview_;
		short magnificationFactorNormalView_;
		int unused2_;		
	};
	struct SCL;
	struct Pane;
	struct Selection;
	struct MergedCells;
	struct LabelRanges;
	struct ConditionalFormattingTable;
	struct HyperlinkTable;
	struct SheetLayout;
	struct SheetProtection;
	struct RangeProtection;

	size_t Read(const char* data);
	size_t Write(char* data);
	size_t DataSize();
	size_t RecordSize();

	BOF bof_;
	Index index_;
	Dimensions dimensions_;
	CellTable cellTable_;
	Window2 window2_;
	YEOF eof_;
};

bool IsRKValueAnInteger(int rkValue);		///< Returns true if the supplied rk value contains an integer. 
bool IsRKValueADouble(int rkValue);			///< Returns true if the supplied rk value contains a double.
double GetDoubleFromRKValue(int rkValue);	///< Convert a rk value to a double.
int GetIntegerFromRKValue(int rkValue);		///< Convert a rk value to an integer.
int GetRKValueFromDouble(double value);		///< Convert a double to a rk value.
int GetRKValueFromInteger(int value);		///< Convert an integer to a rk value.
bool CanStoreAsRKValue(double value);		///< Returns true if the supplied double can be stored as a rk value.

// Forward declarations
class BasicExcel;
class BasicExcelWorksheet;
class BasicExcelCell;

/*******************************************************************************************************/
/*                         Actual classes to read and write to Excel files                             */
/*******************************************************************************************************/
class BasicExcel
{
public:
	BasicExcel();
	BasicExcel(const char* filename);
	~BasicExcel();

public: // File functions.
	void New(int sheets=3);	///< Create a new Excel workbook with a given number of spreadsheets (Minimum 1).
	bool Load(const char* filename);	///< Load an Excel workbook from a file.
	bool Save();	///< Save current Excel workbook to opened file.
	bool SaveAs(const char* filename);	///< Save current Excel workbook to a file.

public: // Worksheet functions.
	size_t GetTotalWorkSheets();	///< Total number of Excel worksheets in current Excel workbook.

	BasicExcelWorksheet* GetWorksheet(size_t sheetIndex);	///< Get a pointer to an Excel worksheet at the given index. Index starts from 0. Returns 0 if index is invalid.
	BasicExcelWorksheet* GetWorksheet(const char* name);	///< Get a pointer to an Excel worksheet that has given ANSI name. Returns 0 if there is no Excel worksheet with the given name.
	BasicExcelWorksheet* GetWorksheet(const wchar_t* name);	///< Get a pointer to an Excel worksheet that has given Unicode name. Returns 0 if there is no Excel worksheet with the given name.
	
	BasicExcelWorksheet* AddWorksheet(int sheetIndex=-1);	///< Add a new Excel worksheet to the given index. Name given to worksheet is SheetX, where X is a number which starts from 1. Index starts from 0. Worksheet is added to the last position if sheetIndex == -1. Returns a pointer to the worksheet if successful, 0 if otherwise.
	BasicExcelWorksheet* AddWorksheet(const char* name, int sheetIndex=-1);	///< Add a new Excel worksheet with given ANSI name to the given index. Index starts from 0. Worksheet is added to the last position if sheetIndex == -1. Returns a pointer to the worksheet if successful, 0 if otherwise.
	BasicExcelWorksheet* AddWorksheet(const wchar_t* name, int sheetIndex=-1);	///< Add a new Excel worksheet with given Unicode name to the given index. Index starts from 0. Worksheet is added to the last position if sheetIndex == -1. Returns a pointer to the worksheet if successful, 0 if otherwise.

	bool DeleteWorksheet(size_t sheetIndex);	///< Delete an Excel worksheet at the given index. Index starts from 0. Returns true if successful, false if otherwise.
	bool DeleteWorksheet(const char* name);		///< Delete an Excel worksheet that has given ANSI name. Returns true if successful, false if otherwise.
	bool DeleteWorksheet(const wchar_t* name);	///< Delete an Excel worksheet that has given Unicode name. Returns true if successful, false if otherwise.

	char* GetAnsiSheetName(size_t sheetIndex);	///< Get the worksheet name at the given index. Index starts from 0. Returns 0 if name is in Unicode format.
	wchar_t* GetUnicodeSheetName(size_t sheetIndex);///< Get the worksheet name at the given index. Index starts from 0. Returns 0 if name is in Ansi format.
	bool GetSheetName(size_t sheetIndex, char* name);		///< Get the worksheet name at the given index. Index starts from 0. Returns false if name is in Unicode format.
	bool GetSheetName(size_t sheetIndex, wchar_t* name);	///< Get the worksheet name at the given index. Index starts from 0. Returns false if name is in Ansi format.

	bool RenameWorksheet(size_t sheetIndex, const char* to);		///< Rename an Excel worksheet at the given index to the given ANSI name. Index starts from 0. Returns true if successful, false if otherwise.
	bool RenameWorksheet(size_t sheetIndex, const wchar_t* to);		///< Rename an Excel worksheet at the given index to the given Unicode name. Index starts from 0. Returns true if successful, false if otherwise.
	bool RenameWorksheet(const char* from, const char* to);			///< Rename an Excel worksheet that has given ANSI name to another ANSI name. Returns true if successful, false if otherwise.
	bool RenameWorksheet(const wchar_t* from, const wchar_t* to);	///< Rename an Excel worksheet that has given Unicode name to another Unicode name. Returns true if successful, false if otherwise.

private: // Functions to read and write raw Excel format.
	size_t Read(const char* data, size_t dataSize);
	size_t Write(char* data);
	void AdjustStreamPositions();
	void AdjustBoundSheetBOFPositions();
	void AdjustDBCellPositions();
	void AdjustExtSSTPositions();

	enum {WORKBOOK_GLOBALS=0x0005, VISUAL_BASIC_MODULE=0x0006,
		  WORKSHEET=0x0010, CHART=0x0020};
	
private: // Internal functions
	void UpdateYExcelWorksheet();	///< Update yesheets_ using information from worksheets_.
	void UpdateWorksheets();		///< Update worksheets_ using information from yesheets_.

public:
	CompoundFile file_;						///< Compound file handler.
	Workbook workbook_;						///< Raw Workbook.
	vector<Worksheet> worksheets_;			///< Raw Worksheets.
	vector<BasicExcelWorksheet> yesheets_;	///< Parsed Worksheets.
};

class BasicExcelWorksheet
{
	friend class BasicExcel;

public:
	BasicExcelWorksheet(BasicExcel* excel, size_t sheetIndex);

public: // Worksheet functions
	char* GetAnsiSheetName();	///< Get the current worksheet name. Returns 0 if name is in Unicode format.
	wchar_t* GetUnicodeSheetName();///< Get the current worksheet name. Returns 0 if name is in Ansi format.
	bool GetSheetName(char* name);		///< Get the current worksheet name. Returns false if name is in Unicode format.
	bool GetSheetName(wchar_t* name);	///< Get the current worksheet name. Returns false if name is in Ansi format.
	bool Rename(const char* to);	///< Rename current Excel worksheet to another ANSI name. Returns true if successful, false if otherwise.
	bool Rename(const wchar_t* to);///< Rename current Excel worksheet to another Unicode name. Returns true if successful, false if otherwise.
	void Print(ostream& os, char delimiter=',', char textQualifier='\0'); ///< Print entire worksheet to an output stream, separating each column with the defined delimiter and enclosing text using the defined textQualifier. Leave out the textQualifier argument if do not wish to have any text qualifiers.

public: // Cell functions
	size_t GetTotalRows();	///< Total number of rows in current Excel worksheet.
	size_t GetTotalCols();	///< Total number of columns in current Excel worksheet.

	BasicExcelCell* Cell(size_t row, size_t col); ///< Return a pointer to an Excel cell. row and col starts from 0. Returns 0 if row exceeds 65535 or col exceeds 255.
	bool EraseCell(size_t row, size_t col); ///< Erase content of a cell. row and col starts from 0. Returns true if successful, false if row or col exceeds range.

private: // Internal functions
	void UpdateCells();	///< Update cells using information from BasicExcel.worksheets_.

private:
	BasicExcel* excel_;					///< Pointer to instance of BasicExcel.
	size_t sheetIndex_;					///< Index of worksheet in workbook.
	size_t maxRows_;					///< Total number of rows in worksheet.
	size_t maxCols_;					///< Total number of columns in worksheet.
	vector<vector<BasicExcelCell> > cells_;	///< Cells matrix.
};

class BasicExcelCell
{
public:
	BasicExcelCell();

public:
	enum {UNDEFINED, INT, DOUBLE, STRING, WSTRING};
	int Type() const;		///< Get type of value stored in current Excel cell. Returns one of the above enums. 

	bool Get(int& val) const;		///< Get an integer value. Returns false if cell does not contain an integer or a double.
	bool Get(double& val) const;	///< Get a double value. Returns false if cell does not contain a double or an integer.
	bool Get(char* str) const;	///< Get an ANSI string. Returns false if cell does not contain an ANSI string.
	bool Get(wchar_t* str) const;	///< Get an Unicode string. Returns false if cell does not contain an Unicode string.
	size_t GetStringLength() const;	///< Return length of ANSI or Unicode string (excluding null character).

	int GetInteger() const;		///< Get an integer value. Returns 0 if cell does not contain an integer.
	double GetDouble() const;		///< Get a double value. Returns 0.0 if cell does not contain a double.
	const char* GetString() const;		///< Get an ANSI string. Returns 0 if cell does not contain an ANSI string.
	const wchar_t* GetWString() const;	///< Get an Unicode string. Returns 0 if cell does not contain an Unicode string.

	friend ostream& operator<<(ostream& os, const BasicExcelCell& cell);	///< Print cell to output stream. Print a null character if cell is undefined.

	void Set(int val);				///< Set content of current Excel cell to an integer.
	void Set(double val);			///< Set content of current Excel cell to a double.
	void Set(const char* str);		///< Set content of current Excel cell to an ANSI string.
	void Set(const wchar_t* str);	///< Set content of current Excel cell to an Unicode string.
	
	void SetInteger(int val);			///< Set content of current Excel cell to an integer.
	void SetDouble(double val);			///< Set content of current Excel cell to a double.
	void SetString(const char* str);	///< Set content of current Excel cell to an ANSI string.
	void SetWString(const wchar_t* str);///< Set content of current Excel cell to an Unicode string.

	void EraseContents();	///< Erase the content of current Excel cell. Set type to UNDEFINED.

private:
	int type_;				///< Type of value stored in current Excel cell. Contains one of the above enums.
	int ival_;				///< Integer value stored in current Excel cell.
	double dval_;			///< Double value stored in current Excel cell.
	vector<char> str_;		///< ANSI string stored in current Excel cell. Include null character.
	vector<wchar_t> wstr_;	///< Unicode string stored in current Excel cell. Include null character.
};

} // Namespace end
#endif

namespace YCompoundFiles
{
/********************************** Start of Class Block *************************************/
// PURPOSE: Manage a file by treating it as blocks of data of a certain size.
Block::Block() : 
	blockSize_(512), fileSize_(0), indexEnd_(0),
	filename_(0) {}

bool Block::Create(const wchar_t* filename)
// PURPOSE: Create a new block file and open it.
// PURPOSE: If file is present, truncate it and then open it.
// PROMISE: Return true if file is successfully created and opened, false if otherwise.
{
	// Create new file
	size_t filenameLength = wcslen(filename);
	char* name = new char[filenameLength+1];
	wcstombs(name, filename, filenameLength);
	name[filenameLength] = 0;

	file_.open(name, ios_base::out | ios_base::trunc);
	file_.close();
	file_.clear();

	// Open the file
	bool ret = this->Open(filename);
	delete[] name;
	return ret;
}

bool Block::Open(const wchar_t* filename, ios_base::openmode mode)
// PURPOSE: Open an existing block file.
// PROMISE: Return true if file is successfully opened, false if otherwise.
{
	// Open existing file for reading or writing or both
	size_t filenameLength = wcslen(filename);
	filename_.resize(filenameLength+1, 0);
	wcstombs(&*(filename_.begin()), filename, filenameLength);

	file_.open(&*(filename_.begin()), mode | ios_base::binary);
	if (!file_.is_open()) return false;

	mode_ = mode;
	
	// Calculate filesize
	if (mode & ios_base::in)
	{
		file_.seekg(0, ios_base::end);
		fileSize_ = file_.tellg();
	}
	else if (mode & ios_base::out)
	{
		file_.seekp(0, ios_base::end);
		fileSize_ = file_.tellp();
	}
	else
	{
		this->Close();
		return false;	
	}
	
	// Calculate last index + 1
	indexEnd_ = fileSize_/blockSize_ + (fileSize_ % blockSize_ ? 1 : 0);
	return true;
}

bool Block::Close()
// PURPOSE: Close the opened block file.
// PROMISE: Return true if file is successfully closed, false if otherwise.
{
	file_.close();
	file_.clear();
	filename_.clear(); 
	fileSize_ = 0; 
	indexEnd_ = 0; 
	blockSize_ = 512;
	return !file_.is_open();
}

bool Block::IsOpen()
// PURPOSE: Check if the block file is still opened.
// PROMISE: Return true if file is still opened, false if otherwise.
{
	return file_.is_open();
}

bool Block::Read(size_t index, char* block)
// PURPOSE: Read a block of data from the opened file at the index position.
// EXPLAIN: index is from [0..].
// PROMISE: Return true if data are successfully read, false if otherwise.
{
	if (!(mode_ & ios_base::in)) return false;
	if (index < indexEnd_)
	{
		file_.seekg(index * blockSize_);
		file_.read(block, blockSize_);
		return !file_.fail();
	}
	else return false;
}

bool Block::Write(size_t index, const char* block)
// PURPOSE: Write a block of data to the opened file at the index position.
// EXPLAIN: index is from [0..].
// PROMISE: Return true if data are successfully written, false if otherwise.
{
	if (!(mode_ & ios_base::out)) return false;
	file_.seekp(index * blockSize_);
	file_.write(block, blockSize_);
	if (indexEnd_ <= index) 
	{
		indexEnd_ = index + 1;
		fileSize_ += blockSize_;
	}
	file_.close();
	file_.clear();
	file_.open(&*(filename_.begin()), mode_ | ios_base::binary);
	return file_.is_open();		
}

bool Block::Swap(size_t index1, size_t index2)
// PURPOSE: Swap two blocks of data in the opened file at the index positions.
// EXPLAIN: index1 and index2 are from [0..].
// PROMISE: Return true if data are successfully swapped, false if otherwise.
{
	if (!(mode_ & ios_base::out)) return false;
	if (index1 < indexEnd_ && index2 < indexEnd_)
	{
		if (index1 == index2) return true;
		
		char* block1 = new char[blockSize_];
		if (!this->Read(index1, block1)) return false;
		
		char* block2 = new char[blockSize_];
		if (!this->Read(index2, block2)) return false;
		
		if (!this->Write(index1, block2)) return false;
		if (!this->Write(index2, block1)) return false;
		
		delete[] block1;
		delete[] block2;
		return true;
	}
	else return false;
}

bool Block::Move(size_t from, size_t to)
// PURPOSE: Move a block of data in the opened file from an index position to another index position.
// EXPLAIN: from and to are from [0..].
// PROMISE: Return true if data are successfully moved, false if otherwise.
{
	if (!(mode_ & ios_base::out)) return false;
	if (from < indexEnd_ && to < indexEnd_)
	{
		if (to > from)
		{
			for (size_t i=from; i!=to; ++i)
			{
				if (!this->Swap(i, i+1)) return false;	
			}
		}
		else
		{
			for (size_t i=from; i!=to; --i)
			{
				if (!this->Swap(i, i-1)) return false;	
			}
		}
		return true;	
	}
	else return false;
}

bool Block::Insert(size_t index, const char* block)
// PURPOSE: Insert a new block of data in the opened file at the index position.
// EXPLAIN: index is from [0..].
// PROMISE: Return true if data are successfully inserted, false if otherwise.
{
	if (!(mode_ & ios_base::out)) return false;
	if (index <= indexEnd_)
	{
		// Write block to end of file
		if (!this->Write(indexEnd_, block)) return false;

		// Move block to index if necessary
		if (index < indexEnd_-1) return this->Move(indexEnd_-1, index);
		else return true;
	}
	else 
	{
		// Write block to index after end of file
		return this->Write(index, block); 
	}
}

bool Block::Erase(size_t index)
// PURPOSE: Erase a block of data in the opened file at the index position.
// EXPLAIN: index is from [0..].
// PROMISE: Return true if data are successfully erased, false if otherwise.
{
	if (!(mode_ & ios_base::out)) return false;
	if (index < indexEnd_)
	{
		fileSize_ -= blockSize_;
		indexEnd_ -= 1;
		
		// Read entire file except the block to be deleted into memory.
		char* buffer = new char[fileSize_];
		for (size_t i=0, j=0; i!=indexEnd_+1; ++i) 
		{
			file_.seekg(i*blockSize_);
			if (i != index) 
			{
				file_.read(buffer+j*blockSize_, blockSize_);
				++j;
			}
		}
		file_.close();
		file_.open(&*(filename_.begin()), ios_base::out | ios_base::trunc | ios_base::binary);
		file_.write(buffer, fileSize_);	// Write the new file.
		file_.close();
		file_.open(&*(filename_.begin()), mode_ | ios_base::binary);
		delete[] buffer;
		return true;
	}
	else return false;
}

bool Block::Erase(vector<size_t>& indices)
// PURPOSE: Erase blocks of data in the opened file at the index positions.
// EXPLAIN: Each index in indices is from [0..].
// PROMISE: Return true if data are successfully erased, false if otherwise.
{
	if (!(mode_ & ios_base::out)) return false;

	// Read entire file except the blocks to be deleted into memory.
	size_t maxIndices = indices.size();
	fileSize_ -= maxIndices*blockSize_;
	char* buffer = new char[fileSize_];
	for (size_t i=0, k=0; i!=indexEnd_; ++i) 
	{
		file_.seekg(i*blockSize_);
		bool toDelete = false;
		for (size_t j=0; j<maxIndices; ++j)
		{
			if (i == indices[j])
			{
				toDelete = true;
				break;
			}
		}
		if (!toDelete) 
		{
			file_.read(buffer+k*blockSize_, blockSize_);
			++k;
		}
	}
	indexEnd_ -= maxIndices;
	
	file_.close();
	file_.open(&*(filename_.begin()), ios_base::out | ios_base::trunc | ios_base::binary);
	file_.write(buffer, fileSize_);	// Write the new file.
	file_.close();
	file_.open(&*(filename_.begin()), mode_ | ios_base::binary);
	delete[] buffer;
	return true;
}
/********************************** End of Class Block ***************************************/

/********************************** Start of Class Header ************************************/
// PURPOSE: Read and write data to a compound file header.
CompoundFile::Header::Header() : 
	fileType_(0xE11AB1A1E011CFD0),
	uk1_(0), uk2_(0), uk3_(0), uk4_(0), uk5_(0x003B), uk6_(0x0003), uk7_(-2),
	log2BigBlockSize_(9), log2SmallBlockSize_(6), 
	uk8_(0), uk9_(0), uk10_(0), uk11_(0x00001000),
	SBATStart_(-2), SBATCount_(0),
	XBATStart_(-2), XBATCount_(0),
	BATCount_(1), propertiesStart_(1) 
{
	BATArray_[0] = 0;	// Initial BAT indices at block 0 (=block 1 in Block)
	fill (BATArray_+1, BATArray_+109, -1);	// Rest of the BATArray is empty
	Initialize();
}

void CompoundFile::Header::Write(char* block)
// PURPOSE: Write header information into a block of data.
// REQUIRE: Block of data must be at least 512 bytes in size.
{
	LittleEndian::Write(block, fileType_, 0x0000, 8);
	LittleEndian::Write(block, uk1_, 0x0008, 4);
	LittleEndian::Write(block, uk2_, 0x000C, 4);
	LittleEndian::Write(block, uk3_, 0x0010, 4);
	LittleEndian::Write(block, uk4_, 0x0014, 4);
	LittleEndian::Write(block, uk5_, 0x0018, 2);
	LittleEndian::Write(block, uk6_, 0x001A, 2);
	LittleEndian::Write(block, uk7_, 0x001C, 2);
	LittleEndian::Write(block, log2BigBlockSize_, 0x001E, 2);
	LittleEndian::Write(block, log2SmallBlockSize_, 0x0020, 4);
	LittleEndian::Write(block, uk8_, 0x0024, 4);
	LittleEndian::Write(block, uk9_, 0x0028, 4);
	LittleEndian::Write(block, BATCount_, 0x002C, 4);
	LittleEndian::Write(block, propertiesStart_, 0x0030, 4);
	LittleEndian::Write(block, uk10_, 0x0034, 4);
	LittleEndian::Write(block, uk11_, 0x0038, 4);
	LittleEndian::Write(block, SBATStart_, 0x003C, 4);
	LittleEndian::Write(block, SBATCount_, 0x0040, 4);
	LittleEndian::Write(block, XBATStart_, 0x0044, 4);
	LittleEndian::Write(block, XBATCount_, 0x0048, 4);
	for (size_t i=0; i<109; ++i) LittleEndian::Write(block, BATArray_[i], 0x004C+i*4, 4);
}

void CompoundFile::Header::Read(char* block)
// PURPOSE: Read header information from a block of data.
// REQUIRE: Block of data must be at least 512 bytes in size.
{
	LittleEndian::Read(block, fileType_, 0x0000, 8);
	LittleEndian::Read(block, uk1_, 0x0008, 4);
	LittleEndian::Read(block, uk2_, 0x000C, 4);
	LittleEndian::Read(block, uk3_, 0x0010, 4);
	LittleEndian::Read(block, uk4_, 0x0014, 4);
	LittleEndian::Read(block, uk5_, 0x0018, 2);
	LittleEndian::Read(block, uk6_, 0x001A, 2);
	LittleEndian::Read(block, uk7_, 0x001C, 2);
	LittleEndian::Read(block, log2BigBlockSize_, 0x001E, 2);
	LittleEndian::Read(block, log2SmallBlockSize_, 0x0020, 4);
	LittleEndian::Read(block, uk8_, 0x0024, 4);
	LittleEndian::Read(block, uk9_, 0x0028, 4);
	LittleEndian::Read(block, BATCount_, 0x002C, 4);
	LittleEndian::Read(block, propertiesStart_, 0x0030, 4);
	LittleEndian::Read(block, uk10_, 0x0034, 4);
	LittleEndian::Read(block, uk11_, 0x0038, 4);
	LittleEndian::Read(block, SBATStart_, 0x003C, 4);
	LittleEndian::Read(block, SBATCount_, 0x0040, 4);
	LittleEndian::Read(block, XBATStart_, 0x0044, 4);
	LittleEndian::Read(block, XBATCount_, 0x0048, 4);
	for (size_t i=0; i<109; ++i) LittleEndian::Read(block, BATArray_[i], 0x004C+i*4, 4);
	Initialize();		
}

void CompoundFile::Header::Initialize()
{
	bigBlockSize_ = (size_t)pow(2.0, log2BigBlockSize_);		// Calculate each big block size.
	smallBlockSize_ = (size_t)pow(2.0, log2SmallBlockSize_);	// Calculate each small block size.
}
/********************************** End of Class Header **************************************/

/********************************** Start of Class Property **********************************/
// PURPOSE: Read and write data to a compound file property.
CompoundFile::Property::Property() :
	nameSize_(0),
	propertyType_(1), nodeColor_(1),
	previousProp_(-1), nextProp_(-1), childProp_(-1),
	uk1_(0), uk2_(0), uk3_(0), uk4_(0), uk5_(0),
	seconds1_(0), days1_(0), seconds2_(0), days2_(0),
	startBlock_(-2), size_(0)
{
	fill (name_, name_+32, 0);
}

void CompoundFile::Property::Write(char* block)
// PURPOSE: Write property information from a block of data.
// REQUIRE: Block of data must be at least 128 bytes in size.
{
	LittleEndian::WriteString(block, name_, 0x00, 32);
	LittleEndian::Write(block, nameSize_, 0x40, 2);
	LittleEndian::Write(block, propertyType_, 0x42, 1);
	LittleEndian::Write(block, nodeColor_, 0x43, 1);
	LittleEndian::Write(block, previousProp_, 0x44, 4);
	LittleEndian::Write(block, nextProp_, 0x48, 4);
	LittleEndian::Write(block, childProp_, 0x4C, 4);
	LittleEndian::Write(block, uk1_, 0x50, 4);
	LittleEndian::Write(block, uk2_, 0x54, 4);
	LittleEndian::Write(block, uk3_, 0x58, 4);
	LittleEndian::Write(block, uk4_, 0x5C, 4);
	LittleEndian::Write(block, uk5_, 0x60, 4);
	LittleEndian::Write(block, seconds1_, 0x64, 4);
	LittleEndian::Write(block, days1_, 0x68, 4);
	LittleEndian::Write(block, seconds2_, 0x6C, 4);
	LittleEndian::Write(block, days2_, 0x70, 4);
	LittleEndian::Write(block, startBlock_, 0x74, 4);
	LittleEndian::Write(block, size_, 0x78, 4);
}

void CompoundFile::Property::Read(char* block)
// PURPOSE: Read property information from a block of data.
// REQUIRE: Block of data must be at least 128 bytes in size.
{
	LittleEndian::ReadString(block, name_, 0x00, 32);
	LittleEndian::Read(block, nameSize_, 0x40, 2);
	LittleEndian::Read(block, propertyType_, 0x42, 1);
	LittleEndian::Read(block, nodeColor_, 0x43, 1);
	LittleEndian::Read(block, previousProp_, 0x44, 4);
	LittleEndian::Read(block, nextProp_, 0x48, 4);
	LittleEndian::Read(block, childProp_, 0x4C, 4);
	LittleEndian::Read(block, uk1_, 0x50, 4);
	LittleEndian::Read(block, uk2_, 0x54, 4);
	LittleEndian::Read(block, uk3_, 0x58, 4);
	LittleEndian::Read(block, uk4_, 0x5C, 4);
	LittleEndian::Read(block, uk5_, 0x60, 4);
	LittleEndian::Read(block, seconds1_, 0x64, 4);
	LittleEndian::Read(block, days1_, 0x68, 4);
	LittleEndian::Read(block, seconds2_, 0x6C, 4);
	LittleEndian::Read(block, days2_, 0x70, 4);
	LittleEndian::Read(block, startBlock_, 0x74, 4);
	LittleEndian::Read(block, size_, 0x78, 4);
}
/********************************** End of Class Property ************************************/

/********************************** Start of Class PropertyTree **********************************/
CompoundFile::PropertyTree::PropertyTree() {};

CompoundFile::PropertyTree::~PropertyTree()
{
	size_t maxChildren = children_.size();
	for (size_t i=0; i<maxChildren; ++i) delete children_[i];
}
/********************************** End of Class PropertyTree ************************************/

/********************************** Start of Class CompoundFile ******************************/
// PURPOSE: Manage a compound file.
CompoundFile::CompoundFile() :
	block_(512), properties_(0), propertyTrees_(0),
	blocksIndices_(0), sblocksIndices_(0) {};

CompoundFile::~CompoundFile() {this->Close();}

/************************* Compound File Functions ***************************/
bool CompoundFile::Create(const wchar_t* filename)
// PURPOSE: Create a new compound file and open it.
// PURPOSE: If file is present, truncate it and then open it.
// PROMISE: Return true if file is successfully created and opened, false if otherwise.
{
	Close();
	file_.Create(filename);

	// Write compound file header
	header_ = Header();
	SaveHeader();

	// Save BAT
	blocksIndices_.clear();
	blocksIndices_.resize(128, -1);
	blocksIndices_[0] = -3;
	blocksIndices_[1] = -2;
	SaveBAT();

	// Save properties
	Property* root = new Property;
	wcscpy(root->name_, L"Root Entry");
	root->propertyType_ = 5;
	properties_.push_back(root);
	SaveProperties();

	// Set property tree
	propertyTrees_ = new PropertyTree;	
	propertyTrees_->parent_ = 0;
	propertyTrees_->self_ = properties_[0];
	propertyTrees_->index_ = 0;
	currentDirectory_ = propertyTrees_;

	return true;
}

bool CompoundFile::Open(const wchar_t* filename, ios_base::openmode mode)
// PURPOSE: Open an existing compound file.
// PROMISE: Return true if file is successfully opened, false if otherwise.
{
	Close();
	if (!file_.Open(filename, mode)) return false;
	
	// Load header
	if (!LoadHeader()) return false;

	// Load BAT information
	LoadBAT();	

	// Load properties
	propertyTrees_ = new PropertyTree;	
	LoadProperties();
	currentDirectory_ = propertyTrees_;

	return true;
}

bool CompoundFile::Close()
// PURPOSE: Close the opened compound file.
// PURPOSE: Reset BAT indices, SBAT indices, properties and properties tree information.
// PROMISE: Return true if file is successfully closed, false if otherwise.
{
	blocksIndices_.clear();
	sblocksIndices_.clear();

	size_t maxProperties = properties_.size();
	for (size_t i=0; i<maxProperties; ++i)
	{
		if (properties_[i]) delete properties_[i];
	}
	properties_.clear();

	if (propertyTrees_) 
	{
		delete propertyTrees_;
		propertyTrees_ = 0;
	}

	previousDirectories_.clear();
	currentDirectory_ = 0;

	return file_.Close();
}

bool CompoundFile::IsOpen()
// PURPOSE: Check if the compound file is still opened.
// PROMISE: Return true if file is still opened, false if otherwise.
{
	return file_.IsOpen();
}

/************************* Directory Functions ***************************/
int CompoundFile::ChangeDirectory(const wchar_t* path)
// PURPOSE: Change to a different directory in the compound file.
// PROMISE: Current directory will not be changed if directory is not present.
{
	previousDirectories_.push_back(currentDirectory_);

	// Handle special cases
	if (wcscmp(path, L".") == 0) 	
	{
		// Current directory
		previousDirectories_.pop_back();
		return SUCCESS;
	}
	if (wcscmp(path, L"..") == 0)
	{
		// Go up 1 directory
		if (currentDirectory_->parent_ != 0)
		{
			currentDirectory_ = currentDirectory_->parent_;
		}
		previousDirectories_.pop_back();
		return SUCCESS;
	}
	if (wcscmp(path, L"\\") == 0)
	{
		// Go to root directory
		currentDirectory_ = propertyTrees_;
		previousDirectories_.pop_back();
		return SUCCESS;
	}

	// Handle normal cases
	size_t ipos = 0;
	size_t npos = 0;
	size_t pathLength = wcslen(path);
	if (pathLength > 0 && path[0] == L'\\')
	{
		// Start from root directory
		currentDirectory_ = propertyTrees_;
		++ipos;
		++npos;
	}
	do
	{
		for (; npos<pathLength; ++npos)
		{
			if (path[npos] == L'\\') break; 
		}

		wchar_t* directory = new wchar_t[npos-ipos+1];
		copy (path+ipos, path+npos, directory);
		directory[npos-ipos] = 0;
		currentDirectory_ = FindProperty(currentDirectory_, directory);
		delete[] directory;
		ipos = npos + 1;
		npos = ipos;
		if (currentDirectory_ == 0)
		{
			// Directory not found
			currentDirectory_ = previousDirectories_.back();
			previousDirectories_.pop_back();
			return DIRECTORY_NOT_FOUND;
		}
	} while (npos < pathLength);
	previousDirectories_.pop_back();
	return SUCCESS;
}

int CompoundFile::MakeDirectory(const wchar_t* path)
// PURPOSE: Create a new directory in the compound file.
// PROMISE: Directory will not be created if it is already present or
// PROMISE: a file with the same name is present.
{
	previousDirectories_.push_back(currentDirectory_);
	Property* property = new Property;
	property->propertyType_ = 1;
	int ret = MakeProperty(path, property);
	currentDirectory_ = previousDirectories_.back();
	previousDirectories_.pop_back();
	SaveHeader();
	SaveBAT();
	SaveProperties();
	return ret;
}

int CompoundFile::PresentWorkingDirectory(wchar_t* path)
// PURPOSE: Get the full path of the current directory in the compound file.
// REQUIRE: path must be large enough to receive the full path information.
{
	previousDirectories_.push_back(currentDirectory_);
	vector<wchar_t> fullpath;
	do
	{
		size_t directoryLength = wcslen(currentDirectory_->self_->name_);
		vector<wchar_t> directory(directoryLength+1);
		directory[0] = L'\\';
		copy (currentDirectory_->self_->name_,
			  currentDirectory_->self_->name_+directoryLength,
			  directory.begin()+1);
		fullpath.insert(fullpath.begin(), directory.begin(), directory.end());		
	} while (currentDirectory_ = currentDirectory_->parent_);

	fullpath.erase(fullpath.begin(), fullpath.begin()+11);
	if (fullpath.empty()) fullpath.push_back(L'\\');
	copy (fullpath.begin(), fullpath.end(), path);
	path[fullpath.size()] = 0;
	currentDirectory_ = previousDirectories_.back();
	previousDirectories_.pop_back();
	return SUCCESS;
}

int CompoundFile::PresentWorkingDirectory(vector<wchar_t>& path)
// PURPOSE: Get the full path of the current directory in the compound file.
{
	previousDirectories_.push_back(currentDirectory_);
	path.clear();
	do
	{
		size_t directoryLength = wcslen(currentDirectory_->self_->name_);
		vector<wchar_t> directory(directoryLength+1);
		directory[0] = L'\\';
		copy (currentDirectory_->self_->name_,
			  currentDirectory_->self_->name_+directoryLength,
			  directory.begin()+1);
		path.insert(path.begin(), directory.begin(), directory.end());		
	} while (currentDirectory_ = currentDirectory_->parent_);

	path.erase(path.begin(), path.begin()+11);
	if (path.empty()) path.push_back(L'\\');
	currentDirectory_ = previousDirectories_.back();
	previousDirectories_.pop_back();
	return SUCCESS;
}

int CompoundFile::RemoveDirectory(const wchar_t* path)
// PURPOSE: Remove a directory in the compound file.
// PROMISE: Directory will not be removed if it has subdirectories or files under it.
{
	PropertyTree* directory = FindProperty(path);
	if (directory == 0) return DIRECTORY_NOT_FOUND;
	if (directory->self_->childProp_ != -1) return DIRECTORY_NOT_EMPTY;
	DeletePropertyTree(directory);
	SaveHeader();
	SaveBAT();
	SaveProperties();
	return SUCCESS;
}

int CompoundFile::DelTree(const wchar_t* path)
// PURPOSE: Remove everything in the path in the compound file, including
// PURPOSE: any files and subdirectories.
{
	previousDirectories_.push_back(currentDirectory_);
	PropertyTree* directory = FindProperty(path);
	if (directory == 0) return DIRECTORY_NOT_FOUND;
	if (directory->self_->childProp_ != -1)
	{
		size_t maxChildren = directory->children_.size();
		wchar_t* curpath = new wchar_t[65535];
		for (size_t i=0; i<maxChildren; ++i)
		{
			currentDirectory_ = directory->children_[i];
			PresentWorkingDirectory(curpath);
			if (directory->children_[i]->self_->propertyType_ == 1)
			{	
				// Directory
				DelTree(curpath);
			}
			else if (directory->children_[i]->self_->propertyType_ == 2)
			{
				// File
				RemoveFile(curpath);
			}
		}
		directory->self_->childProp_ = -1;
		delete[] curpath;
	}

	if (directory->self_->propertyType_ == 1)
	{	
		// Directory
		RemoveDirectory(path);
	}
	else if (directory->self_->propertyType_ == 2)
	{
		// File
		RemoveFile(path);
	}

	currentDirectory_ = previousDirectories_.back();
	previousDirectories_.pop_back();
	return SUCCESS;
}

int CompoundFile::DirectoryList(vector<vector<wchar_t> >& list, const wchar_t* path)
{
	previousDirectories_.push_back(currentDirectory_);
	if (path != 0)
	{
		int ret = ChangeDirectory(path);
		if (ret != SUCCESS) return ret;
	}
	list.clear();
	size_t maxChildren = currentDirectory_->children_.size();
	vector<wchar_t> name(32);
	for (size_t i=0; i<maxChildren; ++i)
	{
		wcscpy(&*(name.begin()), currentDirectory_->children_[i]->self_->name_);
		list.push_back(name);
	}
	currentDirectory_ = previousDirectories_.back();
	previousDirectories_.pop_back();
	return SUCCESS;
}


/************************* File Functions ***************************/
int CompoundFile::MakeFile(const wchar_t* path)
// PURPOSE: Create a new file in the compound file.
// PROMISE: File will not be created if it is already present or
// PROMISE: a directory with the same name is present.
{
	previousDirectories_.push_back(currentDirectory_);
	Property* property = new Property;
	property->propertyType_ = 2;
	int ret = MakeProperty(path, property);
	currentDirectory_ = previousDirectories_.back();
	previousDirectories_.pop_back();
	SaveHeader();
	SaveBAT();
	SaveProperties();
	return ret;
}

int CompoundFile::RemoveFile(const wchar_t* path)
// PURPOSE: Remove a file in the compound file.
{
	int ret = WriteFile(path, 0, 0);
	if (ret == SUCCESS) 
	{
		DeletePropertyTree(FindProperty(path));
		SaveHeader();
		SaveBAT();
		SaveProperties();
		return SUCCESS;
	}
	else return ret;
}

int CompoundFile::FileSize(const wchar_t* path, size_t& size)
// PURPOSE: Get the size of a file in the compound file.
// PROMISE: Return the data size stored in the Root Entry if path = "\".
// PROMISE: size will not be set if file is not present in the compound file.
{
	// Special case of reading root entry
	if (wcscmp(path, L"\\") == 0)
	{
		size = propertyTrees_->self_->size_;
		return SUCCESS;
	}
	
	// Check to see if file is present in the specified directory.
	PropertyTree* property = FindProperty(path);
	if (property == 0) return FILE_NOT_FOUND;
	else
	{
		size = property->self_->size_;
		return SUCCESS;
	}
}

int CompoundFile::ReadFile(const wchar_t* path, char* data)
// PURPOSE: Read a file's data in the compound file.
// REQUIRE: data must be large enough to receive the file's data.
// REQUIRE: The required data size can be obtained by using FileSize().
// PROMISE: Returns the small blocks of data stored by the Root Entry if path = "\".
// PROMISE: data will not be set if file is not present in the compound file.
{
	// Special case of reading root entry
	char* buffer;
	if (wcscmp(path, L"\\") == 0)
	{
		buffer = new char[DataSize(propertyTrees_->self_->startBlock_, true)];
		ReadData(propertyTrees_->self_->startBlock_, buffer, true);
		copy (buffer, buffer+propertyTrees_->self_->size_, data);
		delete[] buffer;
		return SUCCESS;
	}

	// Check to see if file is present in the specified directory.
	PropertyTree* property = FindProperty(path);
	if (property == 0) return FILE_NOT_FOUND;

	if (property->self_->size_ >= 4096)
	{
		// Data stored in normal big blocks
		buffer = new char[DataSize(property->self_->startBlock_, true)];
		ReadData(property->self_->startBlock_, buffer, true);
	}
	else
	{
		// Data stored in small blocks
		buffer = new char[DataSize(property->self_->startBlock_, false)];
		ReadData(property->self_->startBlock_, buffer, false);
	}
	// Truncated the retrieved data to the actual file size.
	copy (buffer, buffer+property->self_->size_, data);
	delete[] buffer;
	return SUCCESS;
}

int CompoundFile::ReadFile(const wchar_t* path, vector<char>& data)
// PURPOSE: Read a file's data in the compound file.
// PROMISE: Returns the small blocks of data stored by the Root Entry if path = "\".
// PROMISE: data will not be set if file is not present in the compound file.
{
	data.clear();
	size_t dataSize;
	int ret = FileSize(path, dataSize);
	if (ret != SUCCESS) return ret;

	data.resize(dataSize);
	return ReadFile(path, &*(data.begin()));
}

int CompoundFile::WriteFile(const wchar_t* path, const char* data, size_t size)
// PURPOSE: Write data to a file in the compound file.
// PROMISE: The file's original data will be replaced by the new data.
{
	PropertyTree* property = FindProperty(path);
	if (property == 0) return FILE_NOT_FOUND;
	
	if (property->self_->size_ >= 4096)
	{
		if (size >= 4096) property->self_->startBlock_ = WriteData(data, size, property->self_->startBlock_, true);
		else
		{
			property->self_->startBlock_ = WriteData(0, 0, property->self_->startBlock_, true);
			property->self_->startBlock_ = WriteData(data, size, property->self_->startBlock_, false);
		}
	}
	else
	{
		if (size < 4096) property->self_->startBlock_ = WriteData(data, size, property->self_->startBlock_, false);
		else
		{
			property->self_->startBlock_ = WriteData(0, 0, property->self_->startBlock_, false);
			property->self_->startBlock_ = WriteData(data, size, property->self_->startBlock_, true);
		}
	}
	property->self_->size_ = size;
	SaveHeader();
	SaveBAT();
	SaveProperties();
	return SUCCESS;
}

int CompoundFile::WriteFile(const wchar_t* path, const vector<char>& data, size_t size)
// PURPOSE: Write data to a file in the compound file.
// PROMISE: The file's original data will be replaced by the new data.
{
	return WriteFile(path, &*(data.begin()), size);
}

/*************ANSI char compound file, directory and file functions******************/
bool CompoundFile::Create(const char* filename)
{
	size_t filenameLength = strlen(filename);
	wchar_t* wname = new wchar_t[filenameLength+1];
	mbstowcs(wname, filename, filenameLength);
	wname[filenameLength] = 0;
	bool ret = Create(wname);
	delete[] wname;
	return ret;
}

bool CompoundFile::Open(const char* filename, ios_base::openmode mode)
{
	size_t filenameLength = strlen(filename);
	wchar_t* wname = new wchar_t[filenameLength+1];
	mbstowcs(wname, filename, filenameLength);
	wname[filenameLength] = 0;
	bool ret = Open(wname, mode);
	delete[] wname;
	return ret;
}

int CompoundFile::ChangeDirectory(const char* path)
{
	size_t pathLength = strlen(path);
	wchar_t* wpath = new wchar_t[pathLength+1];
	mbstowcs(wpath, path, pathLength);
	wpath[pathLength] = 0;
	int ret = ChangeDirectory(wpath);
	delete[] wpath;
	return ret;
}

int CompoundFile::MakeDirectory(const char* path)
{
	size_t pathLength = strlen(path);
	wchar_t* wpath = new wchar_t[pathLength+1];
	mbstowcs(wpath, path, pathLength);
	wpath[pathLength] = 0;
	int ret = MakeDirectory(wpath);
	delete[] wpath;
	return ret;
}

int CompoundFile::PresentWorkingDirectory(char* path)
{
	size_t pathLength = strlen(path);
	wchar_t* wpath = new wchar_t[pathLength+1];
	int ret = PresentWorkingDirectory(wpath);
	if (ret == SUCCESS)
	{
		pathLength = wcslen(wpath);
		wcstombs(path, wpath, pathLength);
		path[pathLength] = 0;
	}
	delete[] wpath;
	return ret;
}

int CompoundFile::PresentWorkingDirectory(vector<char>& path)
{
	vector<wchar_t> wpath;
	int ret = PresentWorkingDirectory(wpath);
	if (ret == SUCCESS)
	{
		size_t pathLength = wpath.size();
		path.resize(pathLength);
		wcstombs(&*(path.begin()), &*(wpath.begin()), pathLength);
		path[pathLength] = 0;
	}
	return ret;
}

int CompoundFile::RemoveDirectory(const char* path)
{
	size_t pathLength = strlen(path);
	wchar_t* wpath = new wchar_t[pathLength+1];
	mbstowcs(wpath, path, pathLength);
	wpath[pathLength] = 0;
	int ret = RemoveDirectory(wpath);
	delete[] wpath;
	return ret;
}

int CompoundFile::DelTree(const char* path)
{
	size_t pathLength = strlen(path);
	wchar_t* wpath = new wchar_t[pathLength+1];
	mbstowcs(wpath, path, pathLength);
	wpath[pathLength] = 0;
	int ret = DelTree(wpath);
	delete[] wpath;
	return ret;
}

int CompoundFile::MakeFile(const char* path)
{
	size_t pathLength = strlen(path);
	wchar_t* wpath = new wchar_t[pathLength+1];
	mbstowcs(wpath, path, pathLength);
	wpath[pathLength] = 0;
	int ret = MakeFile(wpath);
	delete[] wpath;
	return ret;
}

int CompoundFile::RemoveFile(const char* path)
{
	size_t pathLength = strlen(path);
	wchar_t* wpath = new wchar_t[pathLength+1];
	mbstowcs(wpath, path, pathLength);
	wpath[pathLength] = 0;
	int ret = RemoveFile(wpath);
	delete[] wpath;
	return ret;
}

int CompoundFile::FileSize(const char* path, size_t& size)
{
	size_t pathLength = strlen(path);
	wchar_t* wpath = new wchar_t[pathLength+1];
	mbstowcs(wpath, path, pathLength);
	wpath[pathLength] = 0;
	int ret = FileSize(wpath, size);
	delete[] wpath;
	return ret;
}
int CompoundFile::ReadFile(const char* path, char* data)
{
	size_t pathLength = strlen(path);
	wchar_t* wpath = new wchar_t[pathLength+1];
	mbstowcs(wpath, path, pathLength);
	wpath[pathLength] = 0;
	int ret = ReadFile(wpath, data);
	delete[] wpath;
	return ret;
}
int CompoundFile::ReadFile(const char* path, vector<char>& data)
{
	size_t pathLength = strlen(path);
	wchar_t* wpath = new wchar_t[pathLength+1];
	mbstowcs(wpath, path, pathLength);
	wpath[pathLength] = 0;
	int ret = ReadFile(wpath, data);
	delete[] wpath;
	return ret;
}
int CompoundFile::WriteFile(const char* path, char* data, size_t size)
{
	size_t pathLength = strlen(path);
	wchar_t* wpath = new wchar_t[pathLength+1];
	mbstowcs(wpath, path, pathLength);
	wpath[pathLength] = 0;
	int ret = WriteFile(wpath, data, size);
	delete[] wpath;
	return ret;
}
int CompoundFile::WriteFile(const char* path, vector<char>& data, size_t size)
{
	size_t pathLength = strlen(path);
	wchar_t* wpath = new wchar_t[pathLength+1];
	mbstowcs(wpath, path, pathLength);
	wpath[pathLength] = 0;
	int ret = WriteFile(wpath, data, size);
	delete[] wpath;
	return ret;
}

/*********************** Inaccessible General Functions ***************************/
void CompoundFile::IncreaseLocationReferences(vector<size_t> indices)
// PURPOSE: Increase block location references in header, BAT indices and properties,
// PURPOSE: which will be affected by the insertion of new indices contained in indices.
// PROMISE: Block location references which are smaller than all the new indices
// PROMISE: will not be affected.
// PROMISE: SBAT location references will not be affected.
// PROMISE: Changes will not be written to compound file.
{
	size_t maxIndices = indices.size();

	// Change BAT Array references
	{for (size_t i=0; i<109 && header_.BATArray_[i]!=-1; ++i)
	{
		size_t count = 0;
		for (size_t j=0; j<maxIndices; ++j)
		{
			if (header_.BATArray_[i] >= indices[j] &&
				header_.BATArray_[i] != -1) ++count;
		}
		header_.BATArray_[i] += count;	
	}}

	// Change XBAT start block if any
	if (header_.XBATCount_) 
	{
		size_t count = 0;
		for (size_t j=0; j<maxIndices; ++j)
		{
			if (header_.XBATStart_ >= indices[j] &&
				header_.XBATStart_ != -2) ++count;
		}
		header_.XBATStart_ += count;	
	}

	// Change SBAT start block if any
	if (header_.SBATCount_) 
	{
		size_t count = 0;
		for (size_t j=0; j<maxIndices; ++j)
		{
			if (header_.SBATStart_ >= indices[j] &&
				header_.SBATStart_ != -2) ++count;
		}
		header_.SBATStart_ += count;	
	}

	// Change BAT block indices
	size_t maxBATindices = blocksIndices_.size();
	{for (size_t i=0; i<maxBATindices && blocksIndices_[i]!=-1; ++i)
	{
		size_t count = 0;
		for (size_t j=0; j<maxIndices; ++j)
		{
			if (blocksIndices_[i] > indices[j] &&
				blocksIndices_[i] != -2 && 
				blocksIndices_[i] != -3) ++count;
		}
		blocksIndices_[i] += count;	
	}}

	// Change properties start block
	size_t count = 0;
	{for (size_t i=0; i<maxIndices; ++i)
	{
		if (header_.propertiesStart_ >= indices[i] &&
			header_.propertiesStart_ != -2) ++count;
	}}
	header_.propertiesStart_ += count;

	// Change individual properties start block if their size is more than 4096
	size_t maxProperties = properties_.size();
	if (!properties_.empty()) 
	{
		size_t count = 0;
		for (size_t j=0; j<maxIndices; ++j)
		{
			if (properties_[0]->startBlock_ >= indices[j] &&
				properties_[0]->startBlock_ != -2) ++count;
		}
		properties_[0]->startBlock_ += count;	
	}
	{for (size_t i=1; i<maxProperties; ++i)
	{
		if (properties_[i]->size_ >= 4096)
		{
			size_t count = 0;
			for (size_t j=0; j<maxIndices; ++j)
			{
				if (properties_[i]->startBlock_ >= indices[j] &&
					properties_[i]->startBlock_ != -2) ++count;
			}
			properties_[i]->startBlock_ += count;	
		}
	}}
}

void CompoundFile::DecreaseLocationReferences(vector<size_t> indices)
// PURPOSE: Decrease block location references in header, BAT indices and properties,
// PURPOSE: which will be affected by the deletion of indices contained in indices.
// PROMISE: BAT indices pointing to a deleted index will be redirected to point to 
// PROMISE: the location where the deleted index original points to.
// PROMISE: Block location references which are smaller than all the new indices
// PROMISE: will not be affected.
// PROMISE: SBAT location references will not be affected.
// PROMISE: Changes will not be written to compound file.
{
	size_t maxIndices = indices.size();

	// Change BAT Array references
	{for (size_t i=0; i<109 && header_.BATArray_[i]!=-1; ++i)
	{
		size_t count = 0;
		for (size_t j=0; j<maxIndices; ++j)
		{
			if (header_.BATArray_[i] > indices[j] &&
				header_.BATArray_[i] != -1) ++count;
		}
		header_.BATArray_[i] -= count;
	}}

	// Change XBAT start block if any
	if (header_.XBATCount_) 
	{
		size_t count = 0;
		for (size_t j=0; j<maxIndices; ++j)
		{
			if (header_.XBATStart_ > indices[j] &&
				header_.XBATStart_ != -2) ++count;
		}
		header_.XBATStart_ -= count;
	}

	// Change SBAT start block if any
	if (header_.SBATCount_) 
	{
		size_t count = 0;
		for (size_t j=0; j<maxIndices; ++j)
		{
			if (header_.SBATStart_ > indices[j] &&
				header_.SBATStart_ != -2) ++count;
		}
		header_.SBATStart_ -= count;
	}
	
	// Change BAT block indices
	// Redirect BAT indices pointing to a deleted index to point to
	// the location where the deleted index original points to.
	size_t maxBATindices = blocksIndices_.size();
	{for (size_t i=0; i<maxBATindices && blocksIndices_[i]!=-1; ++i)
	{
		bool end;
		do
		{
			end = true;
			for (size_t j=0; j<maxIndices; ++j)
			{
				if (blocksIndices_[i] == indices[j]) 
				{
					blocksIndices_[i] = blocksIndices_[indices[j]];
					end = false;
					break;
				}
			}		
		} while (!end);
	}}
	// Erase indices to be deleted from the block indices
	sort (indices.begin(), indices.end(), greater<size_t>());
	{for (size_t i=0; i<maxIndices; ++i)
	{
		blocksIndices_.erase(blocksIndices_.begin()+indices[i]);
		blocksIndices_.push_back(-1);
	}}

	// Decrease block location references for affected block indices.
	{for (size_t i=0; i<maxBATindices && blocksIndices_[i]!=-1; ++i)
	{
		size_t count = 0;
		for (size_t j=0; j<maxIndices; ++j)
		{
			if (blocksIndices_[i] > indices[j] &&
				blocksIndices_[i] != -2 && 
				blocksIndices_[i] != -3) ++count;
		}
		blocksIndices_[i] -= count;	
	}}

	// Change properties start block
	size_t count = 0;
	{for (size_t i=0; i<maxIndices; ++i)
	{
		if (header_.propertiesStart_ > indices[i] &&
			header_.propertiesStart_ != -2) ++count;
	}}
	header_.propertiesStart_ -= count;

	size_t maxProperties = properties_.size();
	// Change Root Entry start block
	if (!properties_.empty()) 
	{
		size_t count = 0;
		for (size_t j=0; j<maxIndices; ++j)
		{
			if (properties_[0]->startBlock_ > indices[j] &&
				properties_[0]->startBlock_ != -2) ++count;
		}
		properties_[0]->startBlock_ -= count;	
	}
	{for (size_t i=1; i<maxProperties; ++i)
	{
		if (properties_[i]->size_ >= 4096)
		{
			// Change individual properties start block if their size is more than 4096
			size_t count = 0;
			for (size_t j=0; j<maxIndices; ++j)
			{
				if (properties_[i]->startBlock_ > indices[j] &&
					properties_[i]->startBlock_ != -2) ++count;
			}
			properties_[i]->startBlock_ -= count;	
		}
	}}
}

void CompoundFile::SplitPath(const wchar_t* path, 
							 wchar_t*& parentpath, 
							 wchar_t*& propertyname)
// PURPOSE: Get a path's parent path and its name.
// EXPLAIN: E.g. path = "\\Abc\\def\\ghi => parentpath = "\\Abc\\def", propertyname = "ghi".
// REQUIRE: Calling function is responsible for deleting the memory created for
// REQUIRE: parentpath and propertyname.
{
	size_t pathLength = wcslen(path);

	int npos;
	for (npos=pathLength-1; npos>0; --npos)
	{
		if (path[npos] == L'\\') break; 
	}

	if (npos != 0)
	{
		// Get parent path if available
		parentpath = new wchar_t[npos+1];
		copy (path, path+npos, parentpath);
		parentpath[npos] = 0;
		++npos;
	}

	// Get property name (ignore initial "\" if present)
	if (npos==0 && pathLength > 0 && path[0] == L'\\') ++npos;
	propertyname = new wchar_t[pathLength-npos+1];
	copy (path+npos, path+pathLength, propertyname);
	propertyname[pathLength-npos] = 0;	
}

/*********************** Inaccessible Header Functions ***************************/
bool CompoundFile::LoadHeader()
// PURPOSE: Load header information for compound file.
// PROMISE: Return true if file header contain magic number, false if otherwise.
{
	file_.Read(0, &*(block_.begin()));
	header_.Read(&*(block_.begin()));

	// Check magic number to see if it is a compound file 
	if (header_.fileType_ != 0xE11AB1A1E011CFD0) return false;

	block_.resize(header_.bigBlockSize_);		// Resize buffer block
	file_.SetBlockSize(header_.bigBlockSize_);	// Resize block array block size
	return true;
}

void CompoundFile::SaveHeader()
// PURPOSE: Save header information for compound file.
{
	header_.Write(&*(block_.begin()));
	file_.Write(0, &*(block_.begin()));
}

/*********************** Inaccessible BAT Functions ***************************/
void CompoundFile::LoadBAT()
// PURPOSE: Load all block allocation table information for compound file.
{
	// Read BAT indices
	{for (size_t i=0; i<header_.BATCount_; ++i)
	{
		// Load blocksIndices_
		blocksIndices_.resize(blocksIndices_.size()+128, -1);
		file_.Read(header_.BATArray_[i]+1, &*(block_.begin()));
		for (size_t j=0; j<128; ++j) 
		{
			LittleEndian::Read(&*(block_.begin()), blocksIndices_[j+i*128], j*4, 4);
		}
	}}

	// Read XBAT indices
	{for (size_t i=0; i<header_.XBATCount_; ++i)
	{
		blocksIndices_.resize(blocksIndices_.size()+128, -1);
		file_.Read(header_.XBATStart_+i+1, &*(block_.begin()));
		for (size_t j=0; j<128; ++j) 
		{
			LittleEndian::Read(&*(block_.begin()), blocksIndices_[j+((i+109)*128)], j*4, 4);
		}
	}}

	// Read SBAT indices
	{for (size_t i=0; i<header_.SBATCount_; ++i)
	{
		sblocksIndices_.resize(sblocksIndices_.size()+128, -1);
		file_.Read(header_.SBATStart_+i+1, &*(block_.begin()));
		for (size_t j=0; j<128; ++j)
		{
			LittleEndian::Read(&*(block_.begin()), sblocksIndices_[j+i*128], j*4, 4);
		}
	}}
}

void CompoundFile::SaveBAT()
// PURPOSE: Save all block allocation table information for compound file.
{
	// Write BAT indices
	{for (size_t i=0; i<header_.BATCount_; ++i)
	{
		for (size_t j=0; j<128; ++j)
		{
			LittleEndian::Write(&*(block_.begin()), blocksIndices_[j+i*128], j*4, 4);
		}
		file_.Write(header_.BATArray_[i]+1, &*(block_.begin()));
	}}

	// Write XBAT indices
	{for (size_t i=0; i<header_.XBATCount_; ++i)
	{
		for (size_t j=0; j<128; ++j)
		{
			LittleEndian::Write(&*(block_.begin()), blocksIndices_[j+((i+109)*128)], j*4, 4);
		}
		file_.Write(header_.XBATStart_+i+1, &*(block_.begin()));
	}}

	// Write SBAT indices
	{for (size_t i=0; i<header_.SBATCount_; ++i)
	{
		for (size_t j=0; j<128; ++j)
		{
			LittleEndian::Write(&*(block_.begin()), sblocksIndices_[j+i*128], j*4, 4);
		}
		file_.Write(header_.SBATStart_+i+1, &*(block_.begin()));
	}}
}

size_t CompoundFile::DataSize(size_t startIndex, bool isBig)
// PURPOSE: Gets the total size occupied by a property, starting from startIndex.
// EXPLAIN: isBig is true if property uses big blocks, false if it uses small blocks.
// PROMISE: Returns the total size occupied by the property which is the total 
// PROMISE: number of blocks occupied multiply by the block size.
{
	vector<size_t> indices;
	if (isBig) 
	{
		GetBlockIndices(startIndex, indices, true);
		return indices.size()*header_.bigBlockSize_;
	}
	else 
	{
		GetBlockIndices(startIndex, indices, false);
		return indices.size()*header_.smallBlockSize_;
	}	
}

size_t CompoundFile::ReadData(size_t startIndex, char* data, bool isBig)
// PURPOSE: Read a property's data, starting from startIndex.
// REQUIRE: data must be large enough to receive the property's data
// REQUIRE: The required data size can be obtained by using DataSize().
// EXPLAIN: isBig is true if property uses big blocks, false if it uses small blocks.
// PROMISE: Returns the total size occupied by the property which is the total 
// PROMISE: number of blocks occupied multiply by the block size.
{
	vector<size_t> indices;
	if (isBig)
	{
		GetBlockIndices(startIndex, indices, true);
		size_t maxIndices = indices.size();
		for (size_t i=0; i<maxIndices; ++i)
		{
			file_.Read(indices[i]+1, data+i*header_.bigBlockSize_);
		}
		return maxIndices*header_.bigBlockSize_;
	}
	else
	{
		GetBlockIndices(startIndex, indices, false);
		size_t minIndex = *min_element(indices.begin(), indices.end());
		size_t maxIndex = *max_element(indices.begin(), indices.end());
		size_t smallBlocksPerBigBlock = header_.bigBlockSize_ / header_.smallBlockSize_;
		size_t minBlock = minIndex / smallBlocksPerBigBlock;
		size_t maxBlock = maxIndex / smallBlocksPerBigBlock + 
						  (maxIndex % smallBlocksPerBigBlock ? 1 : 0);
		size_t totalBlocks = maxBlock - minBlock;
		char* buffer = new char[DataSize(properties_[0]->startBlock_, true)];
		ReadData(properties_[0]->startBlock_, buffer, true);

		size_t maxIndices = indices.size();
		for (size_t i=0; i<maxIndices; ++i)
		{
			size_t start = (indices[i] - minBlock*smallBlocksPerBigBlock)*header_.smallBlockSize_;
			copy (buffer+start, 
				  buffer+start+header_.smallBlockSize_, 
				  data+i*header_.smallBlockSize_);
		}
		delete[] buffer;
		return maxIndices*header_.smallBlockSize_;
	}
}

size_t CompoundFile::WriteData(const char* data, size_t size, int startIndex, bool isBig)
// PURPOSE: Write data to a property, starting from startIndex.
// EXPLAIN: startIndex can be -2 if property initially has no data.
// EXPLAIN: isBig is true if property uses big blocks, false if it uses small blocks.
// PROMISE: The file's original data will be replaced by the new data.
// PROMISE: Returns the startIndex of new data for the property.
{
	if (isBig)
	{
		if (size==0 && startIndex==-2) return startIndex;

		// Get present indices
		vector<size_t> indices;
		GetBlockIndices(startIndex, indices, true);
		size_t maxPresentBlocks = indices.size();

		// Calculate how many blocks does the data need
		size_t extraSize = size % header_.bigBlockSize_;
		size_t maxNewBlocks = size / header_.bigBlockSize_ + (extraSize ? 1 : 0);

		// Readjust indices and remove blocks if new data size is smaller than original
		int extraBlocks = maxPresentBlocks - maxNewBlocks;
		if (extraBlocks > 0)
		{
			// Place new end marker
			if (maxNewBlocks != 0) blocksIndices_[indices[maxNewBlocks]-1] = -2;
			else startIndex = -2;

			// Get indices of blocks to delete
			vector<size_t> indicesToRemove(extraBlocks);
			copy (indices.begin()+maxNewBlocks, indices.end(), indicesToRemove.begin());
			indices.erase(indices.begin()+maxNewBlocks, indices.end());

			// Remove extra blocks and readjust indices
			FreeBlocks(indicesToRemove, true);
		}

		// Write blocks into available space
		size_t remainingFullBlocks = size / header_.bigBlockSize_;
		size_t curIndex=0;
		if (maxPresentBlocks != 0)
		{
			for (; remainingFullBlocks && curIndex<maxPresentBlocks; 
				   --remainingFullBlocks, ++curIndex)
			{
				file_.Write(indices[curIndex]+1, data+curIndex*header_.bigBlockSize_);			
			} 
		}
		
		// Check if all blocks have been written		
		size_t index;
		if (indices.empty()) index = 0;
		else if (curIndex == 0) index = indices[0];
		else index = (startIndex != -2) ? indices[curIndex-1] : 0;
		if (remainingFullBlocks != 0)
		{			
			// Require extra blocks to write data (i.e. new data is larger than original data
			do
			{
				size_t newIndex = GetFreeBlockIndex(true); // Get new free block to write data
				if (startIndex == -2) startIndex = newIndex; // Get start index
				else LinkBlocks(index, newIndex, true); // Link last index to new index
				file_.Write(newIndex+1, data+curIndex*header_.bigBlockSize_);
				++curIndex;			
				index = newIndex;
			} while (--remainingFullBlocks);
		}
		
		if (extraSize != 0)
		{
			size_t newIndex;
			if (curIndex >= maxPresentBlocks)
			{
				// No more free blocks to write extra block data
				newIndex = GetFreeBlockIndex(true); // Get new free block to write data			
				if (startIndex == -2) startIndex = newIndex;
				else LinkBlocks(index, newIndex,true);
			}
			else newIndex = indices[curIndex];

			// Write extra block after increasing its size to the minimum block size
			vector<char> tempdata(header_.bigBlockSize_, 0);
			copy (data+curIndex*header_.bigBlockSize_, data+curIndex*header_.bigBlockSize_+extraSize, tempdata.begin());
			file_.Write(newIndex+1, &*(tempdata.begin()));			
		}
		return startIndex;
	}
	else
	{
		if (size==0 && startIndex==-2) return startIndex;

		if (size != 0 && properties_[0]->startBlock_ == -2)
		{
			size_t newIndex = GetFreeBlockIndex(true);
			fill (block_.begin(), block_.end(), 0);
			file_.Insert(newIndex, &*(block_.begin()));
			IncreaseLocationReferences(vector<size_t>(1, newIndex));
			properties_[0]->startBlock_ = newIndex;
			properties_[0]->size_ = header_.bigBlockSize_;
		}

		// Get present indices
		vector<size_t> indices;
		GetBlockIndices(startIndex, indices, false);
		size_t maxPresentBlocks = indices.size();

		// Calculate how many blocks does the data need
		size_t extraSize = size % header_.smallBlockSize_;
		size_t maxNewBlocks = size / header_.smallBlockSize_ + (extraSize ? 1 : 0);

		vector<char> smallBlocksData;
		int extraBlocks = maxPresentBlocks - maxNewBlocks;
		if (extraBlocks > 0)
		{
			// Readjust indices and remove blocks
			// Place new end marker
			if (maxNewBlocks != 0) sblocksIndices_[indices[maxNewBlocks]-1] = -2;
			else startIndex = -2;

			// Get indices of blocks to delete
			vector<size_t> indicesToRemove(extraBlocks);
			copy (indices.begin()+maxNewBlocks, indices.end(), indicesToRemove.begin());
			indices.erase(indices.begin()+maxNewBlocks, indices.end());

			// Remove extra blocks and readjust indices
			FreeBlocks(indicesToRemove, false);
		}
		else if (extraBlocks < 0)
		{
			size_t maxBlocks = properties_[0]->size_ / header_.bigBlockSize_ +
						       (properties_[0]->size_ % header_.bigBlockSize_ ? 1 : 0);
			size_t actualSize = maxBlocks * header_.bigBlockSize_;
			smallBlocksData.resize(actualSize);	
			ReadData(properties_[0]->startBlock_, &*(smallBlocksData.begin()), true);
			smallBlocksData.resize(properties_[0]->size_);

			// Readjust indices and add blocks
			size_t newBlocksNeeded = -extraBlocks;
			size_t index = maxPresentBlocks - 1;
			for (size_t i=0; i<newBlocksNeeded; ++i)
			{
				size_t newIndex = GetFreeBlockIndex(false); // Get new free block to write data
				if (startIndex == -2) startIndex = newIndex; // Get start index
				else LinkBlocks(index, newIndex, false);  // Link last index to new index
				smallBlocksData.insert(smallBlocksData.begin()+newIndex,
									   header_.smallBlockSize_, 0);
				index = newIndex;
			}
			properties_[0]->size_ = newBlocksNeeded * header_.smallBlockSize_;
		}
		if (smallBlocksData.empty())
		{
			size_t maxBlocks = properties_[0]->size_ / header_.bigBlockSize_ +
						       (properties_[0]->size_ % header_.bigBlockSize_ ? 1 : 0);
			size_t actualSize = maxBlocks * header_.bigBlockSize_;
			smallBlocksData.resize(actualSize);	
			ReadData(properties_[0]->startBlock_, &*(smallBlocksData.begin()), true);
			smallBlocksData.resize(properties_[0]->size_);
		}

		// Write blocks
		GetBlockIndices(startIndex, indices, false);
		size_t fullBlocks = size / header_.smallBlockSize_;
		for (size_t i=0; i<fullBlocks; ++i)
		{
			copy (data+i*header_.smallBlockSize_,
				  data+i*header_.smallBlockSize_+header_.smallBlockSize_,
				  smallBlocksData.begin()+indices[i]*header_.smallBlockSize_);
		}
		if (extraSize != 0)
		{
			copy (data+fullBlocks*header_.smallBlockSize_,
				  data+fullBlocks*header_.smallBlockSize_+extraSize,
				  smallBlocksData.begin()+indices[fullBlocks]*header_.smallBlockSize_);
		}
		WriteData(&*(smallBlocksData.begin()), properties_[0]->size_,
				 properties_[0]->startBlock_, true);
		return startIndex;
	}
}

void CompoundFile::GetBlockIndices(size_t startIndex, vector<size_t>& indices, bool isBig)
// PURPOSE: Get the indices of blocks where data are stored, starting from startIndex.
// EXPLAIN: isBig is true if property uses big blocks, false if it uses small blocks.
{
	indices.clear();
	if (isBig)
	{
		for (size_t i=startIndex; i!=-2; i=blocksIndices_[i]) indices.push_back(i);
	}
	else 
	{
		for (size_t i=startIndex; i!=-2; i=sblocksIndices_[i]) indices.push_back(i);
	}
}

size_t CompoundFile::GetFreeBlockIndex(bool isBig)
// PURPOSE: Get the index of a new block where data can be stored.
// EXPLAIN: isBig is true if property uses big blocks, false if it uses small blocks.
// PROMISE: It does not physically create a new block in the compound file. 
// PROMISE: It only adjust BAT arrays and indices or SBAT arrays and indices so that
// PROMISE: it gives the index of a new block where data can be inserted.
{
	size_t index;
	if (isBig)
	{
		// Find first free location
		index = distance(blocksIndices_.begin(), 
						 find(blocksIndices_.begin(), 
							  blocksIndices_.end(), -1));
		if (index == blocksIndices_.size())
		{
			ExpandBATArray(true);
			index = distance(blocksIndices_.begin(), 
 							 find(blocksIndices_.begin(), 
								  blocksIndices_.end(), -1));
		}
		blocksIndices_[index] = -2;
	}
	else
	{
		// Find first free location
		index = distance(sblocksIndices_.begin(), 
						 find(sblocksIndices_.begin(), 
							  sblocksIndices_.end(), -1));
		if (index == sblocksIndices_.size())
		{
			ExpandBATArray(false);
			index = distance(sblocksIndices_.begin(), 
							 find(sblocksIndices_.begin(), 
							 	  sblocksIndices_.end(), -1));
		}
		sblocksIndices_[index] = -2;
	}
	return index;
}

void CompoundFile::ExpandBATArray(bool isBig)
// PURPOSE: Create a new block of BAT or SBAT indices.
// EXPLAIN: isBig is true if property uses big blocks, false if it uses small blocks.
{
	size_t newIndex;
	fill (block_.begin(), block_.end(), -1);

	if (isBig)
	{
		size_t BATindex = distance(&header_.BATArray_[0], 
								   find(header_.BATArray_, 
										header_.BATArray_+109, -1));
		if (BATindex < 109)
		{
			// Set new BAT index location
			newIndex = blocksIndices_.size(); // New index location
			file_.Insert(newIndex+1, &*(block_.begin()));
			IncreaseLocationReferences(vector<size_t>(1, newIndex));

			// Update BAT array
			header_.BATArray_[BATindex] = newIndex;
			++header_.BATCount_;
		}
		else
		{
			// No free BAT indices. Increment using XBAT
			// Set new XBAT index location
			if (header_.XBATCount_ != 0)
			{
				newIndex = header_.XBATStart_ + header_.XBATCount_;
				file_.Insert(newIndex, &*(block_.begin()));
				IncreaseLocationReferences(vector<size_t>(1, newIndex));
			}
			else
			{
				newIndex = blocksIndices_.size();
				file_.Insert(newIndex, &*(block_.begin()));
				IncreaseLocationReferences(vector<size_t>(1, newIndex));
				header_.XBATStart_ = newIndex;
			}
			++header_.XBATCount_;
		}
		blocksIndices_.insert(blocksIndices_.begin()+newIndex, -3);
		blocksIndices_.resize(blocksIndices_.size()+127, -1);
	}
	else
	{
		// Set new SBAT index location
		if (header_.SBATCount_ != 0)
		{
			newIndex = header_.SBATStart_ + header_.SBATCount_;
			file_.Insert(newIndex, &*(block_.begin()));
			IncreaseLocationReferences(vector<size_t>(1, newIndex));
		}
		else
		{
			newIndex = GetFreeBlockIndex(true);
			file_.Insert(newIndex, &*(block_.begin()));
			IncreaseLocationReferences(vector<size_t>(1, newIndex));
			header_.SBATStart_ = newIndex;
		}
		++header_.SBATCount_;
		sblocksIndices_.resize(sblocksIndices_.size()+128, -1);
	}
}

void CompoundFile::LinkBlocks(size_t from, size_t to, bool isBig)
// PURPOSE: Link one BAT index to another.
// EXPLAIN: isBig is true if property uses big blocks, false if it uses small blocks.
{
	if (isBig) blocksIndices_[from] = to;	
	else sblocksIndices_[from] = to;
}

void CompoundFile::FreeBlocks(vector<size_t>& indices, bool isBig)
// PURPOSE: Delete blocks of data from compound file.
// EXPLAIN: indices contains indices to blocks of data to be deleted. 
// EXPLAIN: isBig is true if property uses big blocks, false if it uses small blocks.
{
	if (isBig)
	{
		// Decrease all location references before deleting blocks from file.
		DecreaseLocationReferences(indices);
		size_t maxIndices = indices.size();
		{for (size_t i=0; i<maxIndices; ++i) ++indices[i];}	// Increase by 1 because Block index 1 corresponds to index 0 here
		file_.Erase(indices);

		// Shrink BAT indices if necessary
		vector<size_t> indicesToRemove;
		while (distance(find(blocksIndices_.begin(), 
						     blocksIndices_.end(),-1),
						     blocksIndices_.end()) >= 128)
		{			
			blocksIndices_.resize(blocksIndices_.size()-128);
			if (header_.XBATCount_ != 0)
			{
				// Shrink XBAT first
				--header_.XBATCount_;
				indicesToRemove.push_back(header_.XBATStart_+header_.XBATCount_+1); // Add 1 because Block index 1 corresponds to index 0 here
				if (header_.XBATCount_ == 0) header_.XBATStart_ = -2;
			}
			else
			{
				// No XBAT, delete last occupied BAT array element
				size_t BATindex = distance(&header_.BATArray_[0], 
										   find(header_.BATArray_, 
												header_.BATArray_+109, -1));
				if (BATindex != 109)
				{
					--header_.BATCount_;
					indicesToRemove.push_back(header_.BATArray_[BATindex-1]+1); // Add 1 because Block index 1 corresponds to index 0 here
					header_.BATArray_[BATindex-1] = -1;
				}
			}
		}
		// Erase extra BAT indices if present
		if (!indicesToRemove.empty()) file_.Erase(indicesToRemove);
	}
	else
	{
		// Erase block
		size_t maxIndices = indices.size();
		size_t maxBlocks = properties_[0]->size_ / header_.bigBlockSize_ +
						   (properties_[0]->size_ % header_.bigBlockSize_ ? 1 : 0);
		size_t size = maxBlocks * header_.bigBlockSize_;
		char* data = new char[size];
		ReadData(properties_[0]->startBlock_, data, true);
		size_t maxSmallBlocks = properties_[0]->size_ / header_.smallBlockSize_;
		char* newdata = new char[properties_[0]->size_-maxIndices*header_.smallBlockSize_];
		{for (size_t i=0, j=0; i<maxSmallBlocks; ++i)
		{
			if (find(indices.begin(), indices.end(), i) == indices.end()) 
			{
				copy (data+i*header_.smallBlockSize_, 
					  data+i*header_.smallBlockSize_+header_.smallBlockSize_, 
					  newdata+j*header_.smallBlockSize_);
				++j;
			}
		}}
		properties_[0]->startBlock_ = WriteData(newdata, properties_[0]->size_-maxIndices*header_.smallBlockSize_,
											   properties_[0]->startBlock_, true);
		properties_[0]->size_ -= maxIndices*header_.smallBlockSize_;
		delete[] data;
		delete[] newdata;

		// Change SBAT indices
		size_t maxSBATindices = sblocksIndices_.size();
		{for (size_t i=0; i<maxIndices; ++i)
		{
			for (size_t j=0; j<maxSBATindices; ++j)
			{
				if (j == indices[i]) continue;
				if (sblocksIndices_[j] == indices[i]) sblocksIndices_[j] = sblocksIndices_[indices[i]];
				if (sblocksIndices_[j] > indices[i] &&
					sblocksIndices_[j] != -1 &&
					sblocksIndices_[j] != -2) --sblocksIndices_[j];
			}
		}}
		sort (indices.begin(), indices.end(), greater<size_t>());
		{for (size_t i=0; i<maxIndices; ++i)
		{
			sblocksIndices_.erase(sblocksIndices_.begin()+indices[i]);
			sblocksIndices_.push_back(-1);
		}}
		vector<size_t> indicesToRemove;
		while (distance(find(sblocksIndices_.begin(), 
						     sblocksIndices_.end(),-1),
						     sblocksIndices_.end()) >= 128)
		{
			// Shrink SBAT indices if necessary
			sblocksIndices_.resize(sblocksIndices_.size()-128);
			--header_.SBATCount_;
			indicesToRemove.push_back(header_.SBATStart_+header_.SBATCount_);
			if (header_.SBATCount_ == 0) header_.SBATStart_ = -2;
		}
		FreeBlocks(indicesToRemove, true);
	}
}

/*********************** Inaccessible Properties Functions ***************************/
void CompoundFile::LoadProperties()
// PURPOSE: Load properties information for compound file.
{
	// Read properties' data from compound file.
	size_t propertiesSize = DataSize(header_.propertiesStart_, true);
	char* buffer = new char[propertiesSize];
	ReadData(header_.propertiesStart_, buffer, true);

	// Split properties' data into individual property.
	size_t maxPropertiesBlock = propertiesSize / header_.bigBlockSize_;
	size_t propertiesPerBlock = header_.bigBlockSize_ / 128;
	size_t maxProperties = maxPropertiesBlock * propertiesPerBlock;
	size_t maxBlocks = maxProperties / propertiesPerBlock + 
					   (maxProperties % propertiesPerBlock ? 1 : 0);

	for (size_t i=0; i<maxBlocks; ++i)
	{
		for (size_t j=0; j<4; ++j)
		{
			// Read individual property
			Property* property = new Property;
			property->Read(buffer+i*512+j*128);
			if (wcslen(property->name_) == 0)
			{
				delete property;
				break;
			}
			properties_.push_back(property);
		}
	}
	delete[] buffer;

	// Generate property trees
	propertyTrees_->parent_ = 0;
	propertyTrees_->self_ = properties_[0];
	propertyTrees_->index_ = 0;

	InsertPropertyTree(propertyTrees_, 
					   properties_[properties_[0]->childProp_], 
					   properties_[0]->childProp_);
}

void CompoundFile::SaveProperties()
// PURPOSE: Save properties information for compound file.
{
	// Calculate total size required by properties
	size_t maxProperties = properties_.size();
	size_t propertiesPerBlock = header_.bigBlockSize_ / 128;
	size_t maxBlocks = maxProperties / propertiesPerBlock + 
					   (maxProperties % propertiesPerBlock ? 1 : 0);
	size_t propertiesSize = maxBlocks*header_.bigBlockSize_;
	char* buffer = new char[propertiesSize];
	{for (size_t i=0; i<propertiesSize; ++i) buffer[i] = 0;}
	{for (size_t i=0; i<maxProperties; ++i)
	{
		// Save individual property
		properties_[i]->Write(buffer+i*128);
	}}

	// Write properties' data to compound file.
	WriteData(buffer, propertiesSize, header_.propertiesStart_, true);
	delete[] buffer;
}

int CompoundFile::MakeProperty(const wchar_t* path, CompoundFile::Property* property)
// PURPOSE: Create a new property in the compound file.
// EXPLAIN: path is the full path name for the property.
// EXPLAIN: property contains information on the type of property to be created.
{
	wchar_t* parentpath = 0;
	wchar_t* propertyname = 0;
	
	// Change to the specified directory. If specified directory is not present,
	// create it.
	if (wcslen(path) != 0)
	{
		if (path[0] == L'\\') currentDirectory_ = propertyTrees_;
	}
	SplitPath(path, parentpath, propertyname);

	if (propertyname != 0)
	{
		if (parentpath != 0)
		{
			if (ChangeDirectory(parentpath) != SUCCESS)
			{
				int ret = MakeDirectory(parentpath);
				if (ret != SUCCESS) 
				{
					delete[] parentpath;
					delete[] propertyname;
					return ret;
				}
				else ChangeDirectory(parentpath);
			}
			delete[] parentpath;
		}

		// Insert property into specified directory
		size_t propertynameLength = wcslen(propertyname);
		if (propertynameLength >= 32) 
		{
			delete[] propertyname;
			return NAME_TOO_LONG;
		}
		wcscpy(property->name_, propertyname);
		delete[] propertyname;
		property->nameSize_ = propertynameLength*2+2;
		if (FindProperty(currentDirectory_, property->name_) == 0)
		{
			// Find location to insert property
			size_t maxProperties = properties_.size();
			size_t index;
			for (index=1; index<maxProperties; ++index)
			{
				if (*(properties_[index]) > *property) break;
			}
			if (index != maxProperties)
			{
				// Change references for all properties affected by the new property
				IncreasePropertyReferences(propertyTrees_, index);
			}
			properties_.insert(properties_.begin()+index, property);
			InsertPropertyTree(currentDirectory_, property, index);
			return SUCCESS;
		}
		else return DUPLICATE_PROPERTY;
	}
	else 
	{
		if (parentpath != 0) delete[] parentpath;
		return INVALID_PATH;	
	}
}

CompoundFile::PropertyTree* CompoundFile::FindProperty(size_t index)
// PURPOSE: Find property in the compound file, given the index of the property.
// PROMISE: Returns a pointer to the property tree of the property if property
// PROMISE: is present, 0 if otherwise.
{
	if (previousDirectories_.empty()) previousDirectories_.push_back(propertyTrees_);
	PropertyTree* currentTree = previousDirectories_.back();
	if (currentTree->index_ != index)
	{
		size_t maxChildren = currentTree->children_.size();
		for (size_t i=0; i<maxChildren; ++i)
		{
			previousDirectories_.push_back(currentTree->children_[i]);
			PropertyTree* child = FindProperty(index);
			if (child != 0) 
			{
				previousDirectories_.pop_back();
				return child;
			}
		}
	}
	else 
	{
		previousDirectories_.pop_back();
		return currentTree;
	}
	previousDirectories_.pop_back();
	return 0;
}

CompoundFile::PropertyTree* CompoundFile::FindProperty(const wchar_t* path)
// PURPOSE: Find property in the compound file, given the path of the property.
// PROMISE: Returns a pointer to the property tree of the property if property
// PROMISE: is present, 0 if otherwise.
{
	previousDirectories_.push_back(currentDirectory_);

	// Change to specified directory
	wchar_t* parentpath = 0;
	wchar_t* filename = 0;

	if (wcslen(path) != 0)
	{
		if (path[0] == L'\\') currentDirectory_ = propertyTrees_;
	}

	SplitPath(path, parentpath, filename);
	if (parentpath != 0)
	{
		int ret = ChangeDirectory(parentpath);
		delete[] parentpath;
		if (ret != SUCCESS)
		{
			// Cannot change to specified directory
			if (filename != 0) delete[] filename;
			currentDirectory_ = previousDirectories_.back();
			previousDirectories_.pop_back();
			PropertyTree* property = 0;
			return property;
		}
	}

	// Check to see if file is present in the specified directory.
	PropertyTree* property = 0;
	if (filename != 0) 
	{
		property = FindProperty(currentDirectory_, filename);
		delete[] filename;
	}
	currentDirectory_ = previousDirectories_.back();
	previousDirectories_.pop_back();
	return property;	
}

CompoundFile::PropertyTree*
CompoundFile::FindProperty(CompoundFile::PropertyTree* parentTree, 
						   wchar_t* name)
// PURPOSE: Find property in the compound file, given the parent property tree and its name.
// PROMISE: Returns a pointer to the property tree of the property if property
// PROMISE: is present, 0 if otherwise.
{
	if (parentTree->self_->childProp_ != -1)
	{
		size_t maxChildren = parentTree->children_.size();
		for (size_t i=0; i<maxChildren; ++i)
		{
			if (wcscmp(parentTree->children_[i]->self_->name_, name) == 0)
			{
				return parentTree->children_[i];
			}
		}
	}
	return 0;
}

void CompoundFile::InsertPropertyTree(CompoundFile::PropertyTree* parentTree, 
									  CompoundFile::Property* property, 
									  size_t index)
// PURPOSE: Insert a property and all its siblings and children into the property tree.
// REQUIRE: If the property is a new property and its index is already occupied by
// REQUIRE: another property, the calling function has to call IncreasePropertyReferences()
// REQUIRE: first before calling this function.
// EXPLAIN: This function is used by LoadProperty() to initialize the property trees
// EXPLAIN: and MakeProperty() thus resulting in the above requirements.
// EXPLAIN: parentTree is the parent of the new property.
// EXPLAIN: property is the property to be added.
// EXPLAIN: index is the index of the new property.
// PROMISE: The property will be added as the parent tree's child and the parent's 
// PROMISE: child property and all the its children previous property and next property 
// PROMISE: will be readjusted to accomodate the next property.
{
	PropertyTree* tree = new PropertyTree;
	tree->parent_ = parentTree;
	tree->self_ = property;
	tree->index_ = index;

	if (property->previousProp_ != -1)
	{
		InsertPropertyTree(parentTree, 
						   properties_[property->previousProp_], 
						   property->previousProp_);
	}

	if (property->nextProp_ != -1)
	{
		InsertPropertyTree(parentTree, 
						   properties_[property->nextProp_], 
						   property->nextProp_);
	}

	if (property->childProp_ != -1)
	{
		InsertPropertyTree(tree, 
						   properties_[property->childProp_], 
						   property->childProp_);
	}

	// Sort children
	size_t maxChildren = parentTree->children_.size();
	size_t i;
	for (i=0; i<maxChildren; ++i)
	{
		if (index < parentTree->children_[i]->index_) break;
	}
	parentTree->children_.insert(parentTree->children_.begin()+i, tree);

	// Update children indices
	UpdateChildrenIndices(parentTree);
}

void CompoundFile::DeletePropertyTree(CompoundFile::PropertyTree* tree)
// PURPOSE: Delete a property from properties.
// EXPLAIN: tree is the property tree to be deleted.
// PROMISE: The tree's parent's child property and all the its children previous property 
// PROMISE: and next property will be readjusted to accomodate the deleted property.
{
	// Decrease all property references
	DecreasePropertyReferences(propertyTrees_, tree->index_);

	// Remove property
	if (properties_[tree->index_]) delete properties_[tree->index_];
	properties_.erase(properties_.begin()+tree->index_);

	// Remove property from property trees
	size_t maxChildren = tree->parent_->children_.size();
	size_t i;
	for (i=0; i<maxChildren; ++i)
	{
		if (tree->parent_->children_[i]->index_ == tree->index_) break;
	}
	tree->parent_->children_.erase(tree->parent_->children_.begin()+i);

	// Update children indices
	UpdateChildrenIndices(tree->parent_);
}

void CompoundFile::UpdateChildrenIndices(CompoundFile::PropertyTree* parentTree)
{
	// Update indices for 1st to middle child
	size_t maxChildren = parentTree->children_.size();
	if (maxChildren != 0)
	{
		vector<PropertyTree*>& children = parentTree->children_;
		size_t prevChild = 0;
		children[0]->self_->previousProp_ = -1;
		children[0]->self_->nextProp_ = -1;
		size_t curChild;
		for (curChild=1; curChild<=maxChildren/2; ++curChild)
		{
			children[curChild]->self_->previousProp_ = children[prevChild]->index_;
			children[curChild]->self_->nextProp_ = -1;
			prevChild = curChild;
		}

		// Update middle child
		--curChild;
		children[curChild]->parent_->self_->childProp_ = children[curChild]->index_;

		// Update from middle to last child
		size_t nextChild = curChild + 1;
		if (nextChild < maxChildren)
		{
			children[curChild]->self_->nextProp_ = children[nextChild]->index_;
			for (++curChild, ++nextChild;
				 nextChild<maxChildren; 
				 ++curChild, ++nextChild)
			{
				children[curChild]->self_->previousProp_ = -1;
				children[curChild]->self_->nextProp_ = children[nextChild]->index_;
				
			}
 			children[curChild]->self_->previousProp_ = -1;
			children[curChild]->self_->nextProp_ = -1;
		}
	}
	else
	{
		parentTree->self_->childProp_ = -1;
	}
}

void CompoundFile::IncreasePropertyReferences(CompoundFile::PropertyTree* parentTree, 
											  size_t index)
// PURPOSE: Increase all property references (previous property, next property
// PURPOSE: and child property) which will be affected by the insertion of the new index.
// EXPLAIN: The recursive method of going through each property tree is used instead of
// EXPLAIN: using the iterative method of going through each property in properties_ is
// EXPLAIN: because the index in property tree needs to be updated also.
{
	if (parentTree->index_ >= index) ++parentTree->index_;
	if (parentTree->self_->previousProp_ != -1)
	{
		if (parentTree->self_->previousProp_ >= index)
		{
			++parentTree->self_->previousProp_;
		}
	}
	if (parentTree->self_->nextProp_!= -1)
	{
		if (parentTree->self_->nextProp_ >= index)
		{
			++parentTree->self_->nextProp_;
		}
	}
	if (parentTree->self_->childProp_ != -1)
	{
		if (parentTree->self_->childProp_ >= index)
		{
			++parentTree->self_->childProp_;
		}
	}

	size_t maxChildren = parentTree->children_.size();
	for (size_t i=0; i<maxChildren; ++i)
	{
		IncreasePropertyReferences(parentTree->children_[i], index);
	}
}

void CompoundFile::DecreasePropertyReferences(CompoundFile::PropertyTree* parentTree, size_t index)
// PURPOSE: Decrease all property references (previous property, next property
// PURPOSE: and child property) which will be affected by the deletion of the index.
// EXPLAIN: The recursive method of going through each property tree is used instead of
// EXPLAIN: using the iterative method of going through each property in properties_ is
// EXPLAIN: because the index in property tree needs to be updated also.
{
	if (parentTree->index_ > index) --parentTree->index_;
	if (parentTree->self_->previousProp_ != -1)
	{
		if (parentTree->self_->previousProp_ > index)
		{
			--parentTree->self_->previousProp_;
		}
	}
	if (parentTree->self_->nextProp_!= -1)
	{
		if (parentTree->self_->nextProp_ > index)
		{
			--parentTree->self_->nextProp_;
		}
	}
	if (parentTree->self_->childProp_ != -1)
	{
		if (parentTree->self_->childProp_ > index)
		{
			--parentTree->self_->childProp_;
		}
	}

	size_t maxChildren = parentTree->children_.size();
	for (size_t i=0; i<maxChildren; ++i)
	{
		DecreasePropertyReferences(parentTree->children_[i], index);
	}
}
} // YCompoundFiles namespace end

namespace YExcel
{
using namespace YCompoundFiles;
/************************************************************************************************************/
Record::Record() : dataSize_(0), recordSize_(4) {};
Record::~Record() {};
size_t Record::Read(const char* data)
{
	LittleEndian::Read(data, code_, 0, 2);		// Read operation code.
	LittleEndian::Read(data, dataSize_, 2, 2);	// Read size of record.
	data_.assign(data+4, data+4+dataSize_);

	recordSize_ = 4 + dataSize_;

	// Check if next record is a continue record
	continueIndices_.clear();
	short code;
	LittleEndian::Read(data, code, dataSize_+4, 2);
	while (code == CODE::CONTINUE)
	{
		continueIndices_.push_back(dataSize_);

		size_t size;
		LittleEndian::Read(data, size, recordSize_+2, 2);
		data_.insert(data_.end(), data+recordSize_+4, data+recordSize_+4+size);
		dataSize_ += size;
		recordSize_ += 4 + size;

		LittleEndian::Read(data, code, recordSize_, 2);
	};
	return recordSize_;
}
size_t Record::Write(char* data)
{	
	LittleEndian::Write(data, code_, 0, 2);		// Write operation code.
	size_t npos = 2;

	if (continueIndices_.empty())
	{
		size_t size = dataSize_;
		size_t i=0;
		while (size > 8224)
		{
			LittleEndian::Write(data, 8224, npos, 2);	// Write size of record.
			npos += 2;
			size -= 8224;
			copy (data_.begin()+i*8224, data_.begin()+(i+1)*8224, data+npos);
			npos += 8224;

			if (size != 0) 
			{
				++i;
				LittleEndian::Write(data, 0x3C, npos, 2);	// Write CONTINUE code.
				npos += 2;
			}		
		}

		LittleEndian::Write(data, size, npos, 2);	// Write size of record.
		npos += 2;
		copy (data_.begin()+i*8224, data_.begin()+i*8224+size, data+npos);
		npos += size;
	}
	else
	{
		size_t maxContinue = continueIndices_.size();
		size_t size = continueIndices_[0];
		LittleEndian::Write(data, size, npos, 2); // Write size of record
		npos += 2;
		copy (data_.begin(), data_.begin()+size, data+npos);
		npos += size;
		size_t c=0;
		for (c=1; c<maxContinue; ++c)
		{
			LittleEndian::Write(data, 0x3C, npos, 2);	// Write CONTINUE code.
			npos += 2;
			size = continueIndices_[c] - continueIndices_[c-1];
			LittleEndian::Write(data, size, npos, 2);
			npos += 2;
			copy (data_.begin()+continueIndices_[c-1], 
				  data_.begin()+continueIndices_[c],
				  data+npos);
			npos += size;
		}
		LittleEndian::Write(data, 0x3C, npos, 2);	// Write CONTINUE code.
		npos += 2;
		size = data_.size() - continueIndices_[c-1];
		LittleEndian::Write(data, size, npos, 2);
		npos += 2;
		copy (data_.begin()+continueIndices_[c-1], 
			  data_.end(),
			  data+npos);
		npos += size;
	}
	return npos;
}
size_t Record::DataSize() {return dataSize_;}
size_t Record::RecordSize() {return recordSize_;}

/************************************************************************************************************/

/************************************************************************************************************/
BOF::BOF() : Record() {code_ = CODE::BOF; dataSize_ = 16; recordSize_ = 20;}
size_t BOF::Read(const char* data)
{
	Record::Read(data);	
	LittleEndian::Read(data_, version_, 0, 2);
	LittleEndian::Read(data_, type_, 2, 2);
	LittleEndian::Read(data_, buildIdentifier_, 4, 2);
	LittleEndian::Read(data_, buildYear_, 6, 2);
	LittleEndian::Read(data_, fileHistoryFlags_, 8, 4);
	LittleEndian::Read(data_, lowestExcelVersion_, 12, 4);
	return RecordSize();
}
size_t BOF::Write(char* data)
{
	data_.resize(dataSize_);
	LittleEndian::Write(data_, version_, 0, 2);
	LittleEndian::Write(data_, type_, 2, 2);
	LittleEndian::Write(data_, buildIdentifier_, 4, 2);
	LittleEndian::Write(data_, buildYear_, 6, 2);
	LittleEndian::Write(data_, fileHistoryFlags_, 8, 4);
	LittleEndian::Write(data_, lowestExcelVersion_, 12, 4);
	return Record::Write(data);
}
/************************************************************************************************************/

/************************************************************************************************************/
YEOF::YEOF() : Record() {code_ = CODE::YEOF; dataSize_ = 0; recordSize_ = 4;}
/************************************************************************************************************/

/************************************************************************************************************/
SmallString::SmallString() : name_(0), wname_(0) {};
SmallString::~SmallString() {Reset();}
SmallString::SmallString(const SmallString& s) : 
	name_(0), wname_(0), unicode_(s.unicode_)
{
	if (s.name_)
	{
		size_t len = strlen(s.name_);
		name_ = new char[len+1];
		strcpy(name_, s.name_);	
	}
	if (s.wname_)
	{
		size_t len = wcslen(s.wname_);
		wname_ = new wchar_t[len+1];
		wcscpy(wname_, s.wname_);	
	}
}
SmallString& SmallString::operator=(const SmallString& s)
{
	Reset();
	unicode_ = s.unicode_;
	if (s.name_)
	{
		size_t len = strlen(s.name_);
		name_ = new char[len+1];
		strcpy(name_, s.name_);	
	}
	if (s.wname_)
	{
		size_t len = wcslen(s.wname_);
		wname_ = new wchar_t[len+1];
		wcscpy(wname_, s.wname_);	
	}
	return *this;
}
const SmallString& SmallString::operator=(const char* str)
{
	unicode_ = 0;
	Reset();
	size_t len = strlen(str);
	name_ = new char[len+1];
	strcpy(name_, str);
	return *this;
}
const SmallString& SmallString::operator=(const wchar_t* str)
{
	unicode_ = 1;
	Reset();
	size_t len = wcslen(str);
	wname_ = new wchar_t[len+1];
	wcscpy(wname_, str);
	return *this;
}
void SmallString::Reset()
{
	if (name_) {delete[] name_; name_ = 0;}
	if (wname_) {delete[] wname_; wname_ = 0;}
}
size_t SmallString::Read(const char* data)
{
	Reset();
	char stringSize;
	LittleEndian::Read(data, stringSize, 0, 1);
	LittleEndian::Read(data, unicode_, 1, 1);
	size_t bytesRead = 2;
	if (unicode_ == 0)
	{
		// ANSI string	
		name_ = new char[stringSize+1];
		LittleEndian::ReadString(data, name_, 2, stringSize);
		name_[stringSize] = 0;
		bytesRead += stringSize;
	}
	else
	{
		// UNICODE
		wname_ = new wchar_t[stringSize+1];
		LittleEndian::ReadString(data, wname_, 2, stringSize);
		wname_[stringSize] = 0;
		bytesRead += stringSize*2;
	}
	return bytesRead;
}
size_t SmallString::Write(char* data)
{
	size_t stringSize = 0;
	size_t bytesWrite = 0;
	if (unicode_ == 0)
	{
		// ANSI string
		if (name_)
		{
			stringSize = strlen(name_);
			LittleEndian::Write(data, stringSize, 0, 1);
			LittleEndian::Write(data, unicode_, 1, 1);
			LittleEndian::WriteString(data, name_, 2, stringSize);
			bytesWrite = 2 + stringSize;
		}
		else
		{
			LittleEndian::Write(data, stringSize, 0, 1);
			LittleEndian::Write(data, unicode_, 1, 1);
			bytesWrite = 2;
		}
	}
	else
	{
		// UNICODE
		if (wname_)
		{
			stringSize = wcslen(wname_);
			LittleEndian::Write(data, stringSize, 0, 1);
			LittleEndian::Write(data, unicode_, 1, 1);
			LittleEndian::WriteString(data, wname_, 2, stringSize);
			bytesWrite = 2 + stringSize*2;
		}
		else
		{
			LittleEndian::Write(data, stringSize, 0, 1);
			LittleEndian::Write(data, unicode_, 1, 1);
			bytesWrite = 2;
		}
	}
	return bytesWrite;
}
size_t SmallString::DataSize() {return (unicode_ == 0) ? StringSize()+2 : StringSize()*2+2;}
size_t SmallString::RecordSize() {return DataSize();}
size_t SmallString::StringSize()
{
	if (unicode_ == 0)
	{
		if (name_) return strlen(name_);
	}
	else
	{
		if (wname_) return wcslen(wname_);
	}
	return 0;
}
/************************************************************************************************************/

/************************************************************************************************************/
LargeString::LargeString() : unicode_(-1), richtext_(0), phonetic_(0) {};
LargeString::~LargeString() {};
LargeString::LargeString(const LargeString& s) : 
	name_(s.name_), wname_(s.wname_), 
	unicode_(s.unicode_), richtext_(s.richtext_), phonetic_(s.phonetic_) {};
LargeString& LargeString::operator=(const LargeString& s)
{
	unicode_ = s.unicode_;
	richtext_ = s.richtext_;
	phonetic_ = s.phonetic_;
	name_ = s.name_;
	wname_ = s.wname_;
	return *this;
}
const LargeString& LargeString::operator=(const char* str)
{
	unicode_ = 0;
	richtext_ = 0;
	phonetic_ = 0;
	wname_.clear();
	size_t len = strlen(str);
	name_.resize(len+1);
	strcpy(&*(name_.begin()), str);
	return *this;
}
const LargeString& LargeString::operator=(const wchar_t* str)
{
	unicode_ = 1;
	richtext_ = 0;
	phonetic_ = 0;
	name_.clear();
	size_t len = wcslen(str);
	wname_.resize(len+1);
	wcscpy(&*(wname_.begin()), str);	
	return *this;
}
size_t LargeString::Read(const char* data)
{
	size_t stringSize;
	LittleEndian::Read(data, stringSize, 0, 2);
	LittleEndian::Read(data, unicode_, 2, 1);
	size_t npos = 3;
	if (unicode_ & 8) 
	{
		LittleEndian::Read(data, richtext_, npos, 2);
		npos += 2;
	}
	if (unicode_ & 4) LittleEndian::Read(data, phonetic_, npos, 4);
	name_.clear();
	wname_.clear();
	size_t bytesRead = 2;
	if (stringSize>0) bytesRead += ContinueRead(data+2, stringSize);
	else bytesRead = 3;
	return bytesRead;
}
size_t LargeString::ContinueRead(const char* data, size_t size)
{
	if (size == 0) return 0;

	char unicode;
	LittleEndian::Read(data, unicode, 0, 1);
	if (unicode_ == -1) unicode_ = unicode;
	if (unicode_ & 1)
	{
		// Present stored string is uncompressed (16 bit)
		size_t npos = 1;
		if (richtext_) npos += 2;
		if (phonetic_) npos += 4;

		size_t strpos = wname_.size();
		wname_.resize(strpos+size, 0);
		if (unicode & 1)
		{
			LittleEndian::ReadString(data, &*(wname_.begin())+strpos, npos, size);
			npos += size * SIZEOFWCHAR_T;
		}
		else
		{
			// String to be read is in ANSI
			vector<char> name(size);
			LittleEndian::ReadString(data, &*(name.begin()), npos, size);
			mbstowcs(&*(wname_.begin())+strpos, &*(name.begin()), size);
			npos += size;
		}
		if (richtext_) npos += 4*richtext_;
		if (phonetic_) npos += phonetic_;
		return npos;
	}
	else
	{
		// Present stored string has character compression (8 bit)
		size_t npos = 1;
		if (richtext_) npos += 2;
		if (phonetic_) npos += 4;

		size_t strpos = name_.size();
		name_.resize(strpos+size, 0);
		if (unicode & 1)
		{
			// String to be read is in unicode
			vector<wchar_t> name(size);
			LittleEndian::ReadString(data, &*(name.begin()), npos, size);
			wcstombs(&*(name_.begin())+strpos, &*(name.begin()), size);
			npos += size * SIZEOFWCHAR_T;
		}
		else
		{
			LittleEndian::ReadString(data, &*(name_.begin())+strpos, npos, size);
			npos += size;
		}
		if (richtext_) npos += 4*richtext_;
		if (phonetic_) npos += phonetic_;
		return npos;
	}
}
size_t LargeString::Write(char* data)
{
	size_t stringSize = 0;
	size_t bytesWrite = 0;
	if (unicode_ & 1)
	{
		// UNICODE
		unicode_ = 1; // Don't handle richtext or phonetic for now.
		if (!wname_.empty())
		{
			stringSize = wname_.size();
			LittleEndian::Write(data, stringSize, 0, 2);
			LittleEndian::Write(data, unicode_, 2, 1);
			LittleEndian::WriteString(data, &*(wname_.begin()), 3, stringSize);
			bytesWrite = 3 + stringSize * SIZEOFWCHAR_T;
		}
		else
		{
			LittleEndian::Write(data, stringSize, 0, 2);
			LittleEndian::Write(data, unicode_, 2, 1);
			bytesWrite = 3;
		}
	}
	else
	{
		// ANSI string
		unicode_ = 0; // Don't handle richtext or phonetic for now.
		if (!name_.empty())
		{
			stringSize = name_.size();
			LittleEndian::Write(data, stringSize, 0, 2);
			LittleEndian::Write(data, unicode_, 2, 1);
			LittleEndian::WriteString(data, &*(name_.begin()), 3, stringSize);
			bytesWrite = 3 + stringSize;
		}
		else
		{
			LittleEndian::Write(data, stringSize, 0, 2);
			LittleEndian::Write(data, unicode_, 2, 1);
			bytesWrite = 3;
		}
	}
	return bytesWrite;
}
size_t LargeString::DataSize() 
{
	size_t dataSize = StringSize() + 3;
	if (richtext_) dataSize += 2 + 4*richtext_;
	if (phonetic_) dataSize += 4 + phonetic_;
	return dataSize;
}
size_t LargeString::RecordSize() {return DataSize();}
size_t LargeString::StringSize()
{
	if (unicode_ & 1) return wname_.size() * SIZEOFWCHAR_T;
	else return name_.size();
}
/************************************************************************************************************/


/************************************************************************************************************/
Workbook::Workbook()
{
	bof_.version_ = 1536;
	bof_.type_ = 5;
	bof_.buildIdentifier_ = 6560;
	bof_.buildYear_ = 1997;
	bof_.fileHistoryFlags_ = 49353;
	bof_.lowestExcelVersion_ = 774;
}

size_t Workbook::Read(const char* data)
{
	size_t bytesRead = 0;
	short code;
	LittleEndian::Read(data, code, 0, 2);
	while (code != CODE::YEOF)
	{
		switch (code)
		{
			case CODE::BOF:
				bytesRead += bof_.Read(data+bytesRead);
				break;
				
			case CODE::WINDOW1:
				bytesRead += window1_.Read(data+bytesRead);
				break;
				
			case CODE::FONT:
				fonts_.push_back(Font());
				bytesRead += fonts_.back().Read(data+bytesRead);
				break;
				
			case CODE::XF:
				XFs_.push_back(XF());
				bytesRead += XFs_.back().Read(data+bytesRead);
				break;
				
			case CODE::STYLE:
				styles_.push_back(Style());
				bytesRead += styles_.back().Read(data+bytesRead);
				break;
				
			case CODE::BOUNDSHEET:
				boundSheets_.push_back(BoundSheet());
				bytesRead += boundSheets_.back().Read(data+bytesRead);
				break;
				
			case CODE::SST:
				bytesRead += sst_.Read(data+bytesRead);
				break;

//			case CODE::EXTSST:
//				bytesRead += extSST_.Read(data+bytesRead);
//				break;

			default:
				Record rec;
				bytesRead += rec.Read(data+bytesRead);
		}
		LittleEndian::Read(data, code, bytesRead, 2);
	}
	bytesRead += eof_.RecordSize();
	return bytesRead;
}
size_t Workbook::Write(char* data)
{
	size_t bytesWritten = 0;

	bytesWritten += bof_.Write(data+bytesWritten);
	
	bytesWritten += window1_.Write(data+bytesWritten);
	
	size_t maxFonts = fonts_.size();
	{for (size_t i=0; i<maxFonts; ++i) {bytesWritten += fonts_[i].Write(data+bytesWritten);}}
	
	size_t maxXFs = XFs_.size();
	{for (size_t i=0; i<maxXFs; ++i) {bytesWritten += XFs_[i].Write(data+bytesWritten);}}

	size_t maxStyles = styles_.size();
	{for (size_t i=0; i<maxStyles; ++i) {bytesWritten += styles_[i].Write(data+bytesWritten);}}

	size_t maxBoundSheets = boundSheets_.size();
	{for (size_t i=0; i<maxBoundSheets; ++i) {bytesWritten += boundSheets_[i].Write(data+bytesWritten);}}

	bytesWritten += sst_.Write(data+bytesWritten);
//	bytesWritten += extSST_.Write(data+bytesWritten);
	
	bytesWritten += eof_.Write(data+bytesWritten);

	return bytesWritten;
}
size_t Workbook::DataSize() 
{
	size_t size = 0;
	size += bof_.RecordSize();
	size += window1_.RecordSize();

	size_t maxFonts = fonts_.size();
	{for (size_t i=0; i<maxFonts; ++i) {size += fonts_[i].RecordSize();}}
	
	size_t maxXFs = XFs_.size();
	{for (size_t i=0; i<maxXFs; ++i) {size += XFs_[i].RecordSize();}}

	size_t maxStyles = styles_.size();
	{for (size_t i=0; i<maxStyles; ++i) {size += styles_[i].RecordSize();}}

	size_t maxBoundSheets = boundSheets_.size();
	{for (size_t i=0; i<maxBoundSheets; ++i) {size += boundSheets_[i].RecordSize();}}

	size += sst_.RecordSize();
//	size += extSST_.RecordSize();
	size += eof_.RecordSize();
	return size;	
}
size_t Workbook::RecordSize() {return DataSize();}
/************************************************************************************************************/

/************************************************************************************************************/
Workbook::Window1::Window1() : Record(), 
	horizontalPos_(0x78), verticalPos_(0x78), width_(0x3B1F), height_(0x2454),
	options_(0x38), activeWorksheetIndex_(0), firstVisibleTabIndex_(0), selectedWorksheetNo_(1),
	worksheetTabBarWidth_(0x258) {code_ = CODE::WINDOW1; dataSize_ = 18; recordSize_ = 22;}
size_t Workbook::Window1::Read(const char* data)
{
	Record::Read(data);
	LittleEndian::Read(data_, horizontalPos_, 0, 2);
	LittleEndian::Read(data_, verticalPos_, 2, 2);
	LittleEndian::Read(data_, width_, 4, 2);
	LittleEndian::Read(data_, height_, 6, 2);
	LittleEndian::Read(data_, options_, 8, 2);
	LittleEndian::Read(data_, activeWorksheetIndex_, 10, 2);
	LittleEndian::Read(data_, firstVisibleTabIndex_, 12, 2);
	LittleEndian::Read(data_, selectedWorksheetNo_, 14, 2);
	LittleEndian::Read(data_, worksheetTabBarWidth_, 16, 2);
	return RecordSize();
}
size_t Workbook::Window1::Write(char* data)
{
	data_.resize(dataSize_);
	LittleEndian::Write(data_, horizontalPos_, 0, 2);
	LittleEndian::Write(data_, verticalPos_, 2, 2);
	LittleEndian::Write(data_, width_, 4, 2);
	LittleEndian::Write(data_, height_, 6, 2);
	LittleEndian::Write(data_, options_, 8, 2);
	LittleEndian::Write(data_, activeWorksheetIndex_, 10, 2);
	LittleEndian::Write(data_, firstVisibleTabIndex_, 12, 2);
	LittleEndian::Write(data_, selectedWorksheetNo_, 14, 2);
	LittleEndian::Write(data_, worksheetTabBarWidth_, 16, 2);
	return Record::Write(data);
}
/************************************************************************************************************/

/************************************************************************************************************/
Workbook::Font::Font() : Record(), 
	height_(200), options_(0), colourIndex_(0x7FFF), weight_(400), escapementType_(0),
	underlineType_(0), family_(0), characterSet_(0), unused_(0)
{
	code_ = CODE::FONT; 
	dataSize_ = 14;
	recordSize_ = 18;
	name_ = L"Arial";
	name_.unicode_ = 1;
}
size_t Workbook::Font::Read(const char* data)
{
	Record::Read(data);
	LittleEndian::Read(data_, height_, 0, 2);
	LittleEndian::Read(data_, options_, 2, 2);
	LittleEndian::Read(data_, colourIndex_, 4, 2);
	LittleEndian::Read(data_, weight_, 6, 2);
	LittleEndian::Read(data_, escapementType_, 8, 2);
	LittleEndian::Read(data_, underlineType_, 10, 1);
	LittleEndian::Read(data_, family_, 11, 1);
	LittleEndian::Read(data_, characterSet_, 12, 1);
	LittleEndian::Read(data_, unused_, 13, 1);
	name_.Read(&*(data_.begin())+14);
	return RecordSize();
}
size_t Workbook::Font::Write(char* data)
{
	data_.resize(DataSize());
	LittleEndian::Write(data_, height_, 0, 2);
	LittleEndian::Write(data_, options_, 2, 2);
	LittleEndian::Write(data_, colourIndex_, 4, 2);
	LittleEndian::Write(data_, weight_, 6, 2);
	LittleEndian::Write(data_, escapementType_, 8, 2);
	LittleEndian::Write(data_, underlineType_, 10, 1);
	LittleEndian::Write(data_, family_, 11, 1);
	LittleEndian::Write(data_, characterSet_, 12, 1);
	LittleEndian::Write(data_, unused_, 13, 1);
	name_.Write(&*(data_.begin())+14);
	return Record::Write(data);
}
size_t Workbook::Font::DataSize() {return (dataSize_ = 14 + name_.RecordSize());}
size_t Workbook::Font::RecordSize() {return (recordSize_ = DataSize()+4);}
/************************************************************************************************************/

/************************************************************************************************************/
Workbook::XF::XF() : Record(), 
	fontRecordIndex_(0), formatRecordIndex_(0), protectionType_(0xFFF5), alignment_(0x20), rotation_(0x00),
	textProperties_(0x00), usedAttributes_(0x00), borderLines_(0x0000), colour1_(0x0000), colour2_(0x20C0) 
	{code_ = CODE::XF; dataSize_ = 20; recordSize_ = 24;}	
size_t Workbook::XF::Read(const char* data)
{
	Record::Read(data);
	LittleEndian::Read(data_, fontRecordIndex_, 0, 2);
	LittleEndian::Read(data_, formatRecordIndex_, 2, 2);
	LittleEndian::Read(data_, protectionType_, 4, 2);
	LittleEndian::Read(data_, alignment_, 6, 1);
	LittleEndian::Read(data_, rotation_, 7, 1);
	LittleEndian::Read(data_, textProperties_, 8, 1);
	LittleEndian::Read(data_, usedAttributes_, 9, 1);
	LittleEndian::Read(data_, borderLines_, 10, 4);
	LittleEndian::Read(data_, colour1_, 14, 4);
	LittleEndian::Read(data_, colour2_, 18, 2);
	return RecordSize();
}	
size_t Workbook::XF::Write(char* data)
{
	data_.resize(dataSize_);
	LittleEndian::Write(data_, fontRecordIndex_, 0, 2);
	LittleEndian::Write(data_, formatRecordIndex_, 2, 2);
	LittleEndian::Write(data_, protectionType_, 4, 2);
	LittleEndian::Write(data_, alignment_, 6, 1);
	LittleEndian::Write(data_, rotation_, 7, 1);
	LittleEndian::Write(data_, textProperties_, 8, 1);
	LittleEndian::Write(data_, usedAttributes_, 9, 1);
	LittleEndian::Write(data_, borderLines_, 10, 4);
	LittleEndian::Write(data_, colour1_, 14, 4);
	LittleEndian::Write(data_, colour2_, 18, 2);
	return Record::Write(data);
}
/************************************************************************************************************/

/************************************************************************************************************/
Workbook::Style::Style() : Record(),
	XFRecordIndex_(0x8000), identifier_(0), level_(0xFF) 
	{code_ = CODE::STYLE; dataSize_ = 2; recordSize_ = 6;}
size_t Workbook::Style::Read(const char* data)
{
	Record::Read(data);
	LittleEndian::Read(data_, XFRecordIndex_, 0, 2);
	if (XFRecordIndex_ & 0x8000)
	{
		// Built-in styles
		LittleEndian::Read(data_, identifier_, 2, 1);
		LittleEndian::Read(data_, level_, 3, 1);
	}
	else
	{
		// User-defined styles
		name_.Read(&*(data_.begin())+2);
	}
	return RecordSize();
}	
size_t Workbook::Style::Write(char* data)
{
	data_.resize(DataSize());
	LittleEndian::Write(data_, XFRecordIndex_, 0, 2);
	if (XFRecordIndex_ & 0x8000)
	{
		// Built-in styles
		LittleEndian::Write(data_, identifier_, 2, 1);
		LittleEndian::Write(data_, level_, 3, 1);
	}
	else
	{
		// User-defined styles
		name_.Write(&*(data_.begin())+2);
	}
	return Record::Write(data);
}
size_t Workbook::Style::DataSize() {return (dataSize_ = (XFRecordIndex_ & 0x8000) ? 4 : 2+name_.RecordSize());}
size_t Workbook::Style::RecordSize() {return (recordSize_ = DataSize()+4);}
/************************************************************************************************************/

/************************************************************************************************************/
Workbook::BoundSheet::BoundSheet() : Record(),
	BOFpos_(0x0000), visibility_(0), type_(0) 
{
	code_ = CODE::BOUNDSHEET; 
	dataSize_ = 6;
	dataSize_ = 10;
	name_ = "Sheet1";
	name_.unicode_ = false;
}	
size_t Workbook::BoundSheet::Read(const char* data)
{
	Record::Read(data);
	LittleEndian::Read(data_, BOFpos_, 0, 4);
	LittleEndian::Read(data_, visibility_, 4, 1);
	LittleEndian::Read(data_, type_, 5, 1);
	name_.Read(&*(data_.begin())+6);
	return RecordSize();
}	
size_t Workbook::BoundSheet::Write(char* data)
{
	data_.resize(DataSize());
	LittleEndian::Write(data_, BOFpos_, 0, 4);
	LittleEndian::Write(data_, visibility_, 4, 1);
	LittleEndian::Write(data_, type_, 5, 1);
	name_.Write(&*(data_.begin())+6);
	return Record::Write(data);
}
size_t Workbook::BoundSheet::DataSize() {return (dataSize_ = 6+name_.RecordSize());}
size_t Workbook::BoundSheet::RecordSize() {return (recordSize_ = DataSize()+4);}
/************************************************************************************************************/

/************************************************************************************************************/
Workbook::SharedStringTable::SharedStringTable() : Record(),
	stringsTotal_(0), uniqueStringsTotal_(0) {code_ = CODE::SST; dataSize_ = 8; recordSize_ = 12;}
size_t Workbook::SharedStringTable::Read(const char* data)
{
	Record::Read(data);
	LittleEndian::Read(data_, stringsTotal_, 0, 4);
	LittleEndian::Read(data_, uniqueStringsTotal_, 4, 4);
	strings_.clear();
	strings_.resize(uniqueStringsTotal_);
	
	size_t npos = 8;
	if (continueIndices_.empty())
	{
		for (size_t i=0; i<uniqueStringsTotal_; ++i)
		{
			npos += strings_[i].Read(&*(data_.begin())+npos);
		}
	}
	else
	{
		// Require special handling since CONTINUE records are present
		size_t maxContinue = continueIndices_.size();

		for (size_t i=0, c=0; i<uniqueStringsTotal_; ++i)
		{
			char unicode;
			size_t stringSize;
			LittleEndian::Read(data_, stringSize, npos, 2);
			LittleEndian::Read(data_, unicode, npos+2, 1);
			size_t multiplier = unicode & 1 ? 2 : 1;
			if (c >= maxContinue || npos+stringSize*multiplier+3 <= continueIndices_[c])
			{
				// String to be read is not split into two records
				npos += strings_[i].Read(&*(data_.begin())+npos);
			}
			else
			{
				// String to be read is split into two or more records
				int bytesRead = 2;// Start from unicode field

				int size = continueIndices_[c] - npos - 1 - bytesRead;
				++c;
				if (size > 0) 
				{
					size /= multiplier;	// Number of characters available for string in current record.
					bytesRead += strings_[i].ContinueRead(&*(data_.begin())+npos+bytesRead, size);
					stringSize -= size;
					size = 0;
				}
				while (c<maxContinue && npos+stringSize+1>continueIndices_[c])
				{
					size_t dataSize = (continueIndices_[c] - continueIndices_[c-1] - 1) / multiplier;
					bytesRead += strings_[i].ContinueRead(&*(data_.begin())+npos+bytesRead, dataSize);
					stringSize -= dataSize + 1;
					++c;
				};
				if (stringSize>0)
				{
					bytesRead += strings_[i].ContinueRead(&*(data_.begin())+npos+bytesRead, stringSize);
				}
				npos += bytesRead;
			}
		}
	}
	return npos + 4*(npos/8224 + 1);
}	
size_t Workbook::SharedStringTable::Write(char* data)
{
	data_.resize(DataSize());
	LittleEndian::Write(data_, stringsTotal_, 0, 4);
	LittleEndian::Write(data_, uniqueStringsTotal_, 4, 4);

	size_t maxContinue = continueIndices_.size();
	for (size_t i=0, c=0, npos=8; i<uniqueStringsTotal_; ++i)
	{
		npos += strings_[i].Write(&*(data_.begin())+npos);
		if (c<maxContinue && npos==continueIndices_[c]) ++c;
		else if (c<maxContinue && npos > continueIndices_[c])
		{
			// Insert unicode flag where appropriate for CONTINUE records.
			data_.insert(data_.begin()+continueIndices_[c], strings_[i].unicode_);
			data_.pop_back();
			++c;
			++npos;
		}
	}
	return Record::Write(data);
}
size_t Workbook::SharedStringTable::DataSize() 
{
	dataSize_ = 8;
	continueIndices_.clear();
	size_t curMax = 8224;
	for (size_t i=0; i<uniqueStringsTotal_; ++i)
	{
		size_t stringSize = strings_[i].StringSize();
		if (dataSize_+stringSize+3 <= curMax)
		{
			dataSize_ += stringSize + 3;
		}
		else
		{
			// If have >= 12 bytes (2 for size, 1 for unicode and >=9 for data, can split string
			// otherwise, end record and start continue record.
			bool unicode = strings_[i].unicode_ & 1;
			if (curMax - dataSize_ >= 12)
			{
				if (unicode && !((curMax-dataSize_)%2)) --curMax;	// Make sure space reserved for unicode strings is even.
				continueIndices_.push_back(curMax);
				stringSize -= (curMax - dataSize_ - 3);
				dataSize_ = curMax;
				curMax += 8224;

				size_t additionalContinueRecords = unicode ? stringSize/8222 : stringSize/8223; // 8222 or 8223 because the first byte is for unicode identifier
				for (size_t j=0; j<additionalContinueRecords; ++j)
				{
					if (unicode)
					{
						--curMax;
						continueIndices_.push_back(curMax);
						curMax += 8223;
						dataSize_ += 8223;
						stringSize -= 8222;
					}
					else
					{
						continueIndices_.push_back(curMax);
						curMax += 8224;
						dataSize_ += 8224;
						stringSize -= 8223;
					}
				}
				dataSize_ += stringSize + 1;
			}
			else
			{
				continueIndices_.push_back(dataSize_);
				curMax = dataSize_ + 8224;
				if (dataSize_+stringSize+3 < curMax)
				{
					dataSize_ += stringSize + 3;
				}
				else
				{
					// If have >= 12 bytes (2 for size, 1 for unicode and >=9 for data, can split string
					// otherwise, end record and start continue record.
					if (curMax - dataSize_ >= 12)
					{
						if (unicode && !((curMax-dataSize_)%2)) --curMax;	// Make sure space reserved for unicode strings is even.
						continueIndices_.push_back(curMax);
						stringSize -= (curMax - dataSize_ - 3);
						dataSize_ = curMax;
						curMax += 8224;

						size_t additionalContinueRecords = unicode ? stringSize/8222 : stringSize/8223; // 8222 or 8223 because the first byte is for unicode identifier
						for (size_t j=0; j<additionalContinueRecords; ++j)
						{
							if (unicode)
							{
								--curMax;
								continueIndices_.push_back(curMax);
								curMax += 8223;
								dataSize_ += 8223;
								stringSize -= 8222;
							}
							else
							{
								continueIndices_.push_back(curMax);
								curMax += 8224;
								dataSize_ += 8224;
								stringSize -= 8223;
							}
						}
						dataSize_ += stringSize + 1;
					}
				}
			}			
		}		
	}
	return dataSize_;
}
size_t Workbook::SharedStringTable::RecordSize()
{
	size_t dataSize = DataSize();
	return (recordSize_ = dataSize + 4*(dataSize/8224 + 1));
}
/************************************************************************************************************/
Workbook::ExtSST::ExtSST() : Record(),
	stringsTotal_(0), streamPos_(0), firstStringPos_(0), unused_(0) 
{
	code_ = CODE::EXTSST; 
	dataSize_ = 2; 
	recordSize_ = 6;
}

size_t Workbook::ExtSST::Read(const char* data)
{
	Record::Read(data);
	LittleEndian::Read(data_, stringsTotal_, 0, 2);

	size_t maxPortions = (dataSize_-2) / 8;
	streamPos_.clear();
	streamPos_.resize(maxPortions);
	firstStringPos_.clear();
	firstStringPos_.resize(maxPortions);
	unused_.clear();
	unused_.resize(maxPortions);

	for (size_t i=0, npos=2; i<maxPortions; ++i)
	{
		LittleEndian::Read(data_, streamPos_[i], npos, 4);
		LittleEndian::Read(data_, firstStringPos_[i], npos+4, 2);
		LittleEndian::Read(data_, unused_[i], npos+6, 2);
		npos += 8;
	}
	return RecordSize();
}

size_t Workbook::ExtSST::Write(char* data) 
{
	data_.resize(DataSize());
	LittleEndian::Write(data_, stringsTotal_, 0, 2);

	size_t maxPortions = streamPos_.size();
	for (size_t i=0, npos=2; i<maxPortions; ++i)
	{
		LittleEndian::Write(data_, streamPos_[i], npos, 4);
		LittleEndian::Write(data_, firstStringPos_[i], npos+4, 2);
		LittleEndian::Write(data_, unused_[i], npos+6, 2);
		npos += 8;
	}
	return Record::Write(data);
}

size_t Workbook::ExtSST::DataSize()
{
	dataSize_ = 2 + streamPos_.size()*8;
	dataSize_ += (int)(dataSize_/8224)*4;
	return dataSize_;
}

size_t Workbook::ExtSST::RecordSize() {return (recordSize_ = DataSize()+(int)((2+streamPos_.size()*8)/8224)*4)+4;}
/************************************************************************************************************/




/************************************************************************************************************/
Worksheet::Worksheet()
{
	bof_.version_ = 1536;
	bof_.type_ = 16;
	bof_.buildIdentifier_ = 6560;
	bof_.buildYear_ = 1997;
	bof_.fileHistoryFlags_ = 49353;
	bof_.lowestExcelVersion_ = 774;
}

size_t Worksheet::Read(const char* data)
{
	size_t bytesRead = 0;
	short code;
	LittleEndian::Read(data, code, 0, 2);
	while (code != CODE::YEOF)
	{
		switch (code)
		{
			case CODE::BOF:
				bytesRead += bof_.Read(data+bytesRead);
				break;				
				
			case CODE::INDEX:
				bytesRead += index_.Read(data+bytesRead);
				break;
				
			case CODE::DIMENSIONS:
				bytesRead += dimensions_.Read(data+bytesRead);
				break;
				
			case CODE::ROW:
				bytesRead += cellTable_.Read(data+bytesRead);
				break;
				
			case CODE::WINDOW2:
				bytesRead += window2_.Read(data+bytesRead);
				break;

			default:
				Record rec;
				bytesRead += rec.Read(data+bytesRead);
		}
		LittleEndian::Read(data, code, bytesRead, 2);
	}
	bytesRead += eof_.RecordSize();
	return bytesRead;
}
size_t Worksheet::Write(char* data)
{
	size_t bytesWritten = 0;
	bytesWritten += bof_.Write(data+bytesWritten);

	bytesWritten += index_.Write(data+bytesWritten);
	
	bytesWritten += dimensions_.Write(data+bytesWritten);
	
	bytesWritten += cellTable_.Write(data+bytesWritten);

	bytesWritten += window2_.Write(data+bytesWritten);
	
	bytesWritten += eof_.Write(data+bytesWritten);

	return bytesWritten;
}
size_t Worksheet::DataSize() 
{
	size_t dataSize = 0;
	dataSize += bof_.RecordSize();
	dataSize += index_.RecordSize();
	dataSize += dimensions_.RecordSize();
	dataSize += cellTable_.RecordSize();
	dataSize += window2_.RecordSize();
	dataSize += eof_.RecordSize();
	return dataSize;	
}
size_t Worksheet::RecordSize() {return DataSize();}
/************************************************************************************************************/

/************************************************************************************************************/
Worksheet::Index::Index() : Record(), 
	unused1_(0), firstUsedRowIndex_(0), firstUnusedRowIndex_(0), unused2_(0)
	{code_ = CODE::INDEX; dataSize_ = 16; recordSize_ = 20; DBCellPos_.resize(1);}
size_t Worksheet::Index::Read(const char* data)
{
	Record::Read(data);
	LittleEndian::Read(data_, unused1_, 0, 4);
	LittleEndian::Read(data_, firstUsedRowIndex_, 4, 4);
	LittleEndian::Read(data_, firstUnusedRowIndex_, 8, 4);
	LittleEndian::Read(data_, unused2_, 12, 4);
	size_t nm = int(firstUnusedRowIndex_ - firstUsedRowIndex_ - 1) / 32 + 1;
	DBCellPos_.clear();
	DBCellPos_.resize(nm);
	if (dataSize_>16)
	{
		for (size_t i=0; i<nm; ++i)
		{
			LittleEndian::Read(data_, DBCellPos_[i], 16+i*4, 4);
		}
	}
	return RecordSize();
}	
size_t Worksheet::Index::Write(char* data)
{
	data_.resize(DataSize());
	LittleEndian::Write(data_, unused1_, 0, 4);
	LittleEndian::Write(data_, firstUsedRowIndex_, 4, 4);
	LittleEndian::Write(data_, firstUnusedRowIndex_, 8, 4);
	LittleEndian::Write(data_, unused2_, 12, 4);
	size_t nm = DBCellPos_.size();
	for (size_t i=0; i<nm; ++i)
	{
		LittleEndian::Write(data_, DBCellPos_[i], 16+i*4, 4);
	}
	return Record::Write(data);
}
size_t Worksheet::Index::DataSize() {return (dataSize_ = 16 + DBCellPos_.size()*4);}
size_t Worksheet::Index::RecordSize() 
{
	size_t dataSize = DataSize();
	return (recordSize_ = dataSize + 4*(dataSize/8224 + 1));
}

/************************************************************************************************************/

/************************************************************************************************************/
Worksheet::Dimensions::Dimensions() : Record(),
	firstUsedRowIndex_(0), lastUsedRowIndexPlusOne_(0), 
	firstUsedColIndex_(0), lastUsedColIndexPlusOne_(0),
	unused_(0) {code_ = CODE::DIMENSIONS; dataSize_ = 14; recordSize_ = 18;}
size_t Worksheet::Dimensions::Read(const char* data)
{
	Record::Read(data);
	LittleEndian::Read(data_, firstUsedRowIndex_, 0, 4);
	LittleEndian::Read(data_, lastUsedRowIndexPlusOne_, 4, 4);
	LittleEndian::Read(data_, firstUsedColIndex_, 8, 2);
	LittleEndian::Read(data_, lastUsedColIndexPlusOne_, 10, 2);
	LittleEndian::Read(data_, unused_, 12, 2);
	return RecordSize();
}	
size_t Worksheet::Dimensions::Write(char* data)
{
	data_.resize(DataSize());
	LittleEndian::Write(data_, firstUsedRowIndex_, 0, 4);
	LittleEndian::Write(data_, lastUsedRowIndexPlusOne_, 4, 4);
	LittleEndian::Write(data_, firstUsedColIndex_, 8, 2);
	LittleEndian::Write(data_, lastUsedColIndexPlusOne_, 10, 2);
	LittleEndian::Write(data_, unused_, 12, 2);
	return Record::Write(data);
}
/************************************************************************************************************/

/************************************************************************************************************/
Worksheet::CellTable::RowBlock::CellBlock::Blank::Blank() : Record(), 
	rowIndex_(0), colIndex_(0), XFRecordIndex_(0) {code_ = CODE::BLANK; dataSize_ = 6; recordSize_ = 10;}
size_t Worksheet::CellTable::RowBlock::CellBlock::Blank::Read(const char* data)
{
	Record::Read(data);
	LittleEndian::Read(data_, rowIndex_, 0, 2);
	LittleEndian::Read(data_, colIndex_, 2, 2);
	LittleEndian::Read(data_, XFRecordIndex_, 4, 2);
	return RecordSize();
}	
size_t Worksheet::CellTable::RowBlock::CellBlock::Blank::Write(char* data)
{
	data_.resize(DataSize());
	LittleEndian::Write(data_, rowIndex_, 0, 2);
	LittleEndian::Write(data_, colIndex_, 2, 2);
	LittleEndian::Write(data_, XFRecordIndex_, 4, 2);
	return Record::Write(data);
}

Worksheet::CellTable::RowBlock::CellBlock::BoolErr::BoolErr() : Record(),
	rowIndex_(0), colIndex_(0), XFRecordIndex_(0), value_(0), error_(0)
	{code_ = CODE::BOOLERR; dataSize_ = 8; recordSize_ = 12;}
size_t Worksheet::CellTable::RowBlock::CellBlock::BoolErr::Read(const char* data)
{
	Record::Read(data);
	LittleEndian::Read(data_, rowIndex_, 0, 2);
	LittleEndian::Read(data_, colIndex_, 2, 2);
	LittleEndian::Read(data_, XFRecordIndex_, 4, 2);
	LittleEndian::Read(data_, value_, 6, 1);
	LittleEndian::Read(data_, error_, 7, 1);
	return RecordSize();
}	
size_t Worksheet::CellTable::RowBlock::CellBlock::BoolErr::Write(char* data)
{
	data_.resize(DataSize());
	LittleEndian::Write(data_, rowIndex_, 0, 2);
	LittleEndian::Write(data_, colIndex_, 2, 2);
	LittleEndian::Write(data_, XFRecordIndex_, 4, 2);
	LittleEndian::Write(data_, value_, 6, 1);
	LittleEndian::Write(data_, error_, 7, 1);
	return Record::Write(data);
}

Worksheet::CellTable::RowBlock::CellBlock::LabelSST::LabelSST() : Record(),
	rowIndex_(0), colIndex_(0), XFRecordIndex_(0), SSTRecordIndex_(0)
	{code_ = CODE::LABELSST; dataSize_ = 10; recordSize_ = 14;}
size_t Worksheet::CellTable::RowBlock::CellBlock::LabelSST::Read(const char* data)
{
	Record::Read(data);
	LittleEndian::Read(data_, rowIndex_, 0, 2);
	LittleEndian::Read(data_, colIndex_, 2, 2);
	LittleEndian::Read(data_, XFRecordIndex_, 4, 2);
	LittleEndian::Read(data_, SSTRecordIndex_, 6, 4);
	return RecordSize();
}	
size_t Worksheet::CellTable::RowBlock::CellBlock::LabelSST::Write(char* data)
{
	data_.resize(DataSize());
	LittleEndian::Write(data_, rowIndex_, 0, 2);
	LittleEndian::Write(data_, colIndex_, 2, 2);
	LittleEndian::Write(data_, XFRecordIndex_, 4, 2);
	LittleEndian::Write(data_, SSTRecordIndex_, 6, 4);
	return Record::Write(data);
}

Worksheet::CellTable::RowBlock::CellBlock::MulBlank::MulBlank() : Record(),
	rowIndex_(0), firstColIndex_(0), lastColIndex_(0)
	{code_ = CODE::MULBLANK; dataSize_ = 10; recordSize_ = 14;}
size_t Worksheet::CellTable::RowBlock::CellBlock::MulBlank::Read(const char* data)
{
	Record::Read(data);
	LittleEndian::Read(data_, rowIndex_, 0, 2);
	LittleEndian::Read(data_, firstColIndex_, 2, 2);
	LittleEndian::Read(data_, lastColIndex_, dataSize_-2, 2);
	size_t nc = lastColIndex_ - firstColIndex_ + 1; 
	XFRecordIndices_.clear();
	XFRecordIndices_.resize(nc);
	for (size_t i=0; i<nc; ++i)
	{
		LittleEndian::Read(data_, XFRecordIndices_[i], 4+i*2, 2);	
	}
	return RecordSize();
}	
size_t Worksheet::CellTable::RowBlock::CellBlock::MulBlank::Write(char* data)
{
	data_.resize(DataSize());
	LittleEndian::Write(data_, rowIndex_, 0, 2);
	LittleEndian::Write(data_, firstColIndex_, 2, 2);
	LittleEndian::Write(data_, lastColIndex_, dataSize_-2, 2);
	size_t nc = XFRecordIndices_.size();
	for (size_t i=0; i<nc; ++i)
	{
		LittleEndian::Write(data_, XFRecordIndices_[i], 4+i*2, 2);	
	}
	return Record::Write(data);
}
size_t Worksheet::CellTable::RowBlock::CellBlock::MulBlank::DataSize() {return (dataSize_ = 6 + XFRecordIndices_.size()*2);}
size_t Worksheet::CellTable::RowBlock::CellBlock::MulBlank::RecordSize() 
{
	size_t dataSize = DataSize();
	return (recordSize_ = dataSize + 4*(dataSize/8224 + 1));
}

Worksheet::CellTable::RowBlock::CellBlock::MulRK::XFRK::XFRK() :
	XFRecordIndex_(0), RKValue_(0) {};
void Worksheet::CellTable::RowBlock::CellBlock::MulRK::XFRK::Read(const char* data)
{
	LittleEndian::Read(data, XFRecordIndex_, 0, 2);
	LittleEndian::Read(data, RKValue_, 2, 4);
}	
void Worksheet::CellTable::RowBlock::CellBlock::MulRK::XFRK::Write(char* data)
{
	LittleEndian::Write(data, XFRecordIndex_, 0, 2);
	LittleEndian::Write(data, RKValue_, 2, 4);
}

Worksheet::CellTable::RowBlock::CellBlock::MulRK::MulRK() : Record(),
	rowIndex_(0), firstColIndex_(0), lastColIndex_(0) {code_ = CODE::MULRK; dataSize_ = 10; recordSize_ = 14;}
size_t Worksheet::CellTable::RowBlock::CellBlock::MulRK::Read(const char* data)
{
	Record::Read(data);
	LittleEndian::Read(data_, rowIndex_, 0, 2);
	LittleEndian::Read(data_, firstColIndex_, 2, 2);
	LittleEndian::Read(data_, lastColIndex_, dataSize_-2, 2);
	size_t nc = lastColIndex_ - firstColIndex_ + 1; 
	XFRK_.clear();
	XFRK_.resize(nc);
	for (size_t i=0; i<nc; ++i)
	{
		XFRK_[i].Read(&*(data_.begin())+4+i*6);
	}
	return RecordSize();
}	
size_t Worksheet::CellTable::RowBlock::CellBlock::MulRK::Write(char* data)
{
	data_.resize(DataSize());
	LittleEndian::Write(data_, rowIndex_, 0, 2);
	LittleEndian::Write(data_, firstColIndex_, 2, 2);
	LittleEndian::Write(data_, lastColIndex_, dataSize_-2, 2);
	size_t nc = XFRK_.size();
	for (size_t i=0; i<nc; ++i)
	{
		XFRK_[i].Write(&*(data_.begin())+4+i*6);
	}
	return Record::Write(data);
}
size_t Worksheet::CellTable::RowBlock::CellBlock::MulRK::DataSize() {return (dataSize_ = 6 + XFRK_.size()*6);}
size_t Worksheet::CellTable::RowBlock::CellBlock::MulRK::RecordSize() 
{
	size_t dataSize = DataSize();
	return (recordSize_ = dataSize + 4*(dataSize/8224 + 1));
}

Worksheet::CellTable::RowBlock::CellBlock::Number::Number() : Record(),
	rowIndex_(0), colIndex_(0), XFRecordIndex_(0), value_(0) {code_ = CODE::NUMBER; dataSize_ = 14; recordSize_ = 18;}
size_t Worksheet::CellTable::RowBlock::CellBlock::Number::Read(const char* data)
{
	Record::Read(data);
	LittleEndian::Read(data_, rowIndex_, 0, 2);
	LittleEndian::Read(data_, colIndex_, 2, 2);
	LittleEndian::Read(data_, XFRecordIndex_, 4, 2);
	__int64 value;
	LittleEndian::Read(data_, value, 6, 8);
	intdouble_.intvalue_ = value;
	value_ = intdouble_.doublevalue_;
	return RecordSize();
}	
size_t Worksheet::CellTable::RowBlock::CellBlock::Number::Write(char* data)
{
	data_.resize(DataSize());
	LittleEndian::Write(data_, rowIndex_, 0, 2);
	LittleEndian::Write(data_, colIndex_, 2, 2);
	LittleEndian::Write(data_, XFRecordIndex_, 4, 2);
	intdouble_.doublevalue_ = value_;
	__int64 value = intdouble_.intvalue_;
	LittleEndian::Write(data_, value, 6, 8);
	return Record::Write(data);
}

Worksheet::CellTable::RowBlock::CellBlock::RK::RK() : Record(),
	rowIndex_(0), colIndex_(0), XFRecordIndex_(0), value_(0) {code_ = CODE::RK; dataSize_ = 10; recordSize_ = 14;}
size_t Worksheet::CellTable::RowBlock::CellBlock::RK::Read(const char* data)
{
	Record::Read(data);
	LittleEndian::Read(data_, rowIndex_, 0, 2);
	LittleEndian::Read(data_, colIndex_, 2, 2);
	LittleEndian::Read(data_, XFRecordIndex_, 4, 2);
	LittleEndian::Read(data_, value_, 6, 4);
	return RecordSize();
}	
size_t Worksheet::CellTable::RowBlock::CellBlock::RK::Write(char* data)
{
	data_.resize(DataSize());
	LittleEndian::Write(data_, rowIndex_, 0, 2);
	LittleEndian::Write(data_, colIndex_, 2, 2);
	LittleEndian::Write(data_, XFRecordIndex_, 4, 2);
	LittleEndian::Write(data_, value_, 6, 4);
	return Record::Write(data);
}

Worksheet::CellTable::RowBlock::CellBlock::Formula::Formula() : Record(),
	rowIndex_(0), colIndex_(0), XFRecordIndex_(0), options_(0), unused_(0), type_(-1)
	{code_ = CODE::FORMULA; dataSize_ = 18; recordSize_ = 22;}
size_t Worksheet::CellTable::RowBlock::CellBlock::Formula::Read(const char* data)
{
	Record::Read(data);
	LittleEndian::Read(data_, rowIndex_, 0, 2);
	LittleEndian::Read(data_, colIndex_, 2, 2);
	LittleEndian::Read(data_, XFRecordIndex_, 4, 2);
	LittleEndian::ReadString(data_, result_, 6, 8);
	LittleEndian::Read(data_, options_, 14, 2);
	LittleEndian::Read(data_, unused_, 16, 2);
	RPNtoken_.clear();
	RPNtoken_.resize(dataSize_-18);
	LittleEndian::ReadString(data_, &*(RPNtoken_.begin()), 18, dataSize_-18);

	size_t offset = dataSize_ + 4;
	short code;
	LittleEndian::Read(data, code, offset, 2);
	switch (code)
	{
		case CODE::ARRAY:
			type_ = code;
			array_.Read(data+offset);
			offset += array_.RecordSize();
			break;
			
		case CODE::SHRFMLA:	
			type_ = code;
			shrfmla_.Read(data+offset);
			offset += shrfmla_.RecordSize();
			break;
			
		case CODE::SHRFMLA1:	
			type_ = code;
			shrfmla1_.Read(data+offset);
			offset += shrfmla1_.RecordSize();
			break;

		case CODE::TABLE:	
			type_ = code;
			table_.Read(data+offset);
			offset += table_.RecordSize();
			break;
	}
	LittleEndian::Read(data, code, offset, 2);
	if (code == CODE::STRING) string_.Read(data+offset);
	return RecordSize();
}	
size_t Worksheet::CellTable::RowBlock::CellBlock::Formula::Write(char* data)
{
	data_.resize(DataSize());
	LittleEndian::Write(data_, rowIndex_, 0, 2);
	LittleEndian::Write(data_, colIndex_, 2, 2);
	LittleEndian::Write(data_, XFRecordIndex_, 4, 2);
	LittleEndian::WriteString(data_, result_, 6, 8);
	LittleEndian::Write(data_, options_, 14, 2);
	LittleEndian::Write(data_, unused_, 16, 2);
	LittleEndian::WriteString(data_, &*(RPNtoken_.begin()), 18, RPNtoken_.size());
	Record::Write(data);

	size_t offset = dataSize_ + 4;
	switch (type_)
	{
		case CODE::ARRAY:
			array_.Write(data+offset);
			offset += array_.RecordSize();
			break;
			
		case CODE::SHRFMLA:	
			shrfmla_.Write(data+offset);
			offset += shrfmla_.RecordSize();
			break;
			
		case CODE::SHRFMLA1:	
			shrfmla1_.Write(data+offset);
			offset += shrfmla1_.RecordSize();
			break;

		case CODE::TABLE:	
			table_.Write(data+offset);
			offset += table_.RecordSize();
			break;
	}
	if (string_.DataSize() != 0) string_.Write(data+offset);
	return RecordSize();
}
size_t Worksheet::CellTable::RowBlock::CellBlock::Formula::DataSize() {return (dataSize_ = 18 + RPNtoken_.size());}
size_t Worksheet::CellTable::RowBlock::CellBlock::Formula::RecordSize() 
{
	size_t dataSize = DataSize();
	recordSize_ = dataSize + 4*(dataSize/8224 + 1);
	
	switch (type_)
	{
		case CODE::ARRAY:
			recordSize_ += array_.RecordSize();
			break;
			
		case CODE::SHRFMLA:	
			recordSize_ += shrfmla_.RecordSize();
			break;
			
		case CODE::SHRFMLA1:	
			recordSize_ += shrfmla1_.RecordSize();
			break;

		case CODE::TABLE:	
			recordSize_ += table_.RecordSize();
			break;
	}
	if (string_.DataSize() != 0) recordSize_ += string_.RecordSize();
	return (recordSize_);
}

Worksheet::CellTable::RowBlock::CellBlock::Formula::Array::Array() : Record(),
	firstRowIndex_(0), lastRowIndex_(0), firstColIndex_(0), lastColIndex_(0),
	options_(0), unused_(0)
	{code_ = CODE::ARRAY; dataSize_ = 12; recordSize_ = 16;}
size_t Worksheet::CellTable::RowBlock::CellBlock::Formula::Array::Read(const char* data)
{
	Record::Read(data);
	LittleEndian::Read(data_, firstRowIndex_, 0, 2);
	LittleEndian::Read(data_, lastRowIndex_, 2, 2);
	LittleEndian::Read(data_, firstColIndex_, 4, 1);
	LittleEndian::Read(data_, lastColIndex_, 5, 1);
	LittleEndian::Read(data_, options_, 6, 2);
	LittleEndian::Read(data_, unused_, 8, 4);
	formula_.clear();
	formula_.resize(dataSize_-12);
	LittleEndian::ReadString(data_, &*(formula_.begin()), 12, dataSize_-12);
	return RecordSize();
}	
size_t Worksheet::CellTable::RowBlock::CellBlock::Formula::Array::Write(char* data)
{
	data_.resize(DataSize());
	LittleEndian::Write(data_, firstRowIndex_, 0, 2);
	LittleEndian::Write(data_, lastRowIndex_, 2, 2);
	LittleEndian::Write(data_, firstColIndex_, 4, 1);
	LittleEndian::Write(data_, lastColIndex_, 5, 1);
	LittleEndian::Write(data_, options_, 6, 2);
	LittleEndian::Write(data_, unused_, 8, 4);
	LittleEndian::WriteString(data_, &*(formula_.begin()), 12, formula_.size());
	return Record::Write(data);
}
size_t Worksheet::CellTable::RowBlock::CellBlock::Formula::Array::DataSize() {return (dataSize_ = 12 + formula_.size());}
size_t Worksheet::CellTable::RowBlock::CellBlock::Formula::Array::RecordSize() 
{
	size_t dataSize = DataSize();
	return (recordSize_ = dataSize + 4*(dataSize/8224 + 1));
}

Worksheet::CellTable::RowBlock::CellBlock::Formula::ShrFmla::ShrFmla() : Record(),
	firstRowIndex_(0), lastRowIndex_(0), firstColIndex_(0), lastColIndex_(0),
	unused_(0)
	{code_ = CODE::SHRFMLA; dataSize_ = 8; recordSize_ = 12;}
size_t Worksheet::CellTable::RowBlock::CellBlock::Formula::ShrFmla::Read(const char* data)
{
	Record::Read(data);
	LittleEndian::Read(data_, firstRowIndex_, 0, 2);
	LittleEndian::Read(data_, lastRowIndex_, 2, 2);
	LittleEndian::Read(data_, firstColIndex_, 4, 1);
	LittleEndian::Read(data_, lastColIndex_, 5, 1);
	LittleEndian::Read(data_, unused_, 6, 2);
	formula_.clear();
	formula_.resize(dataSize_-8);
	LittleEndian::ReadString(data_, &*(formula_.begin()), 8, dataSize_-8);
	return RecordSize();
}	
size_t Worksheet::CellTable::RowBlock::CellBlock::Formula::ShrFmla::Write(char* data)
{
	data_.resize(DataSize());
	LittleEndian::Write(data_, firstRowIndex_, 0, 2);
	LittleEndian::Write(data_, lastRowIndex_, 2, 2);
	LittleEndian::Write(data_, firstColIndex_, 4, 1);
	LittleEndian::Write(data_, lastColIndex_, 5, 1);
	LittleEndian::Write(data_, unused_, 6, 2);
	LittleEndian::WriteString(data_, &*(formula_.begin()), 8, formula_.size());
	return Record::Write(data);
}
size_t Worksheet::CellTable::RowBlock::CellBlock::Formula::ShrFmla::DataSize() {return (dataSize_ = 8 + formula_.size());}
size_t Worksheet::CellTable::RowBlock::CellBlock::Formula::ShrFmla::RecordSize() 
{
	size_t dataSize = DataSize();
	return (recordSize_ = dataSize + 4*(dataSize/8224 + 1));
}

Worksheet::CellTable::RowBlock::CellBlock::Formula::ShrFmla1::ShrFmla1() : Record(),
	firstRowIndex_(0), lastRowIndex_(0), firstColIndex_(0), lastColIndex_(0),
	unused_(0)
	{code_ = CODE::SHRFMLA1; dataSize_ = 8; recordSize_ = 12;}
size_t Worksheet::CellTable::RowBlock::CellBlock::Formula::ShrFmla1::Read(const char* data)
{
	Record::Read(data);
	LittleEndian::Read(data_, firstRowIndex_, 0, 2);
	LittleEndian::Read(data_, lastRowIndex_, 2, 2);
	LittleEndian::Read(data_, firstColIndex_, 4, 1);
	LittleEndian::Read(data_, lastColIndex_, 5, 1);
	LittleEndian::Read(data_, unused_, 6, 2);
	formula_.clear();
	formula_.resize(dataSize_-8);
	LittleEndian::ReadString(data_, &*(formula_.begin()), 8, dataSize_-8);
	return RecordSize();
}	
size_t Worksheet::CellTable::RowBlock::CellBlock::Formula::ShrFmla1::Write(char* data)
{
	data_.resize(DataSize());
	LittleEndian::Write(data_, firstRowIndex_, 0, 2);
	LittleEndian::Write(data_, lastRowIndex_, 2, 2);
	LittleEndian::Write(data_, firstColIndex_, 4, 1);
	LittleEndian::Write(data_, lastColIndex_, 5, 1);
	LittleEndian::Write(data_, unused_, 6, 2);
	LittleEndian::WriteString(data_, &*(formula_.begin()), 8, formula_.size());
	return Record::Write(data);
}
size_t Worksheet::CellTable::RowBlock::CellBlock::Formula::ShrFmla1::DataSize() {return (dataSize_ = 8 + formula_.size());}
size_t Worksheet::CellTable::RowBlock::CellBlock::Formula::ShrFmla1::RecordSize() 
{
	size_t dataSize = DataSize();
	return (recordSize_ = dataSize + 4*(dataSize/8224 + 1));
}

Worksheet::CellTable::RowBlock::CellBlock::Formula::Table::Table() : Record(),
	firstRowIndex_(0), lastRowIndex_(0), firstColIndex_(0), lastColIndex_(0), options_(0), 
	inputCellRowIndex_(0), inputCellColIndex_(0), 
	inputCellColumnInputRowIndex_(0), inputCellColumnInputColIndex_(0)
	{code_ = CODE::TABLE; dataSize_ = 16; recordSize_ = 20;}
size_t Worksheet::CellTable::RowBlock::CellBlock::Formula::Table::Read(const char* data)
{
	Record::Read(data);
	LittleEndian::Read(data_, firstRowIndex_, 0, 2);
	LittleEndian::Read(data_, lastRowIndex_, 2, 2);
	LittleEndian::Read(data_, firstColIndex_, 4, 1);
	LittleEndian::Read(data_, lastColIndex_, 5, 1);
	LittleEndian::Read(data_, options_, 6, 2);
	LittleEndian::Read(data_, inputCellRowIndex_, 8, 2);
	LittleEndian::Read(data_, inputCellColIndex_, 10, 2);
	LittleEndian::Read(data_, inputCellColumnInputRowIndex_, 12, 2);
	LittleEndian::Read(data_, inputCellColumnInputColIndex_, 14, 2);
	return RecordSize();
}	
size_t Worksheet::CellTable::RowBlock::CellBlock::Formula::Table::Write(char* data)
{
	data_.resize(DataSize());
	LittleEndian::Write(data_, firstRowIndex_, 0, 2);
	LittleEndian::Write(data_, lastRowIndex_, 2, 2);
	LittleEndian::Write(data_, firstColIndex_, 4, 1);
	LittleEndian::Write(data_, lastColIndex_, 5, 1);
	LittleEndian::Write(data_, options_, 6, 2);
	LittleEndian::Write(data_, inputCellRowIndex_, 8, 2);
	LittleEndian::Write(data_, inputCellColIndex_, 10, 2);
	LittleEndian::Write(data_, inputCellColumnInputRowIndex_, 12, 2);
	LittleEndian::Write(data_, inputCellColumnInputColIndex_, 14, 2);
	return Record::Write(data);
}

Worksheet::CellTable::RowBlock::CellBlock::Formula::String::String() : Record()
	{code_ = CODE::STRING; dataSize_ = 0; recordSize_ = 4;}
size_t Worksheet::CellTable::RowBlock::CellBlock::Formula::String::Read(const char* data)
{
	Record::Read(data);
	string_.clear();
	string_.resize(dataSize_);
	LittleEndian::ReadString(data_, &*(string_.begin()), 0, dataSize_);
	return RecordSize();
}	
size_t Worksheet::CellTable::RowBlock::CellBlock::Formula::String::Write(char* data)
{
	data_.resize(DataSize());
	LittleEndian::WriteString(data_, &*(string_.begin()), 0, string_.size());
	return Record::Write(data);
}
size_t Worksheet::CellTable::RowBlock::CellBlock::Formula::String::DataSize() {return (dataSize_ = string_.size());}
size_t Worksheet::CellTable::RowBlock::CellBlock::Formula::String::RecordSize() 
{
	size_t dataSize = DataSize();
	return (recordSize_ = dataSize + 4*(dataSize/8224 + 1));
}
/************************************************************************************************************/

/************************************************************************************************************/
Worksheet::CellTable::RowBlock::Row::Row() : Record(),
	rowIndex_(0), firstCellColIndex_(0), lastCellColIndexPlusOne_(0), height_(255),
	unused1_(0), unused2_(0), options_(0) {code_ = CODE::ROW; dataSize_ = 16; recordSize_ = 20;}
size_t Worksheet::CellTable::RowBlock::Row::Read(const char* data)
{
	Record::Read(data);
	LittleEndian::Read(data_, rowIndex_, 0, 2);
	LittleEndian::Read(data_, firstCellColIndex_, 2, 2);
	LittleEndian::Read(data_, lastCellColIndexPlusOne_, 4, 2);
	LittleEndian::Read(data_, height_, 6, 2);
	LittleEndian::Read(data_, unused1_, 8, 2);
	LittleEndian::Read(data_, unused2_, 10, 2);
	LittleEndian::Read(data_, options_, 12, 4);
	return RecordSize();
}	
size_t Worksheet::CellTable::RowBlock::Row::Write(char* data)
{
	data_.resize(DataSize());
	LittleEndian::Write(data_, rowIndex_, 0, 2);
	LittleEndian::Write(data_, firstCellColIndex_, 2, 2);
	LittleEndian::Write(data_, lastCellColIndexPlusOne_, 4, 2);
	LittleEndian::Write(data_, height_, 6, 2);
	LittleEndian::Write(data_, unused1_, 8, 2);
	LittleEndian::Write(data_, unused2_, 10, 2);
	LittleEndian::Write(data_, options_, 12, 4);
	return Record::Write(data);
}
/************************************************************************************************************/

/************************************************************************************************************/
Worksheet::CellTable::RowBlock::CellBlock::CellBlock() : 
	type_(-1), normalType_(true) {};
Worksheet::CellTable::RowBlock::CellBlock::~CellBlock() {};
size_t Worksheet::CellTable::RowBlock::CellBlock::Read(const char* data)
{
	size_t bytesRead = 0;
	LittleEndian::Read(data, type_, 0, 2);
	switch (type_)
	{
		case CODE::BLANK:
			bytesRead += blank_.Read(data);
			break;
			
		case CODE::BOOLERR:	
			bytesRead += boolerr_.Read(data);
			break;
			
		case CODE::LABELSST:	
			bytesRead += labelsst_.Read(data);
			break;
			
		case CODE::MULBLANK:	
			bytesRead += mulblank_.Read(data);
			break;
			
		case CODE::MULRK:	
			bytesRead += mulrk_.Read(data);
			break;
			
		case CODE::NUMBER:	
			bytesRead += number_.Read(data);
			break;
			
		case CODE::RK:	
			bytesRead += rk_.Read(data);
			break;

		case CODE::FORMULA:
			bytesRead += formula_.Read(data);
			break;
	}
	return bytesRead;
}	
size_t Worksheet::CellTable::RowBlock::CellBlock::Write(char* data)
{
	size_t bytesWritten = 0;
	switch (type_)
	{
		case CODE::BLANK:
			bytesWritten += blank_.Write(data);
			break;
			
		case CODE::BOOLERR:	
			bytesWritten += boolerr_.Write(data);
			break;
			
		case CODE::LABELSST:	
			bytesWritten += labelsst_.Write(data);
			break;
			
		case CODE::MULBLANK:	
			bytesWritten += mulblank_.Write(data);
			break;
			
		case CODE::MULRK:	
			bytesWritten += mulrk_.Write(data);
			break;
			
		case CODE::NUMBER:	
			bytesWritten += number_.Write(data);
			break;
			
		case CODE::RK:	
			bytesWritten += rk_.Write(data);
			break;

		case CODE::FORMULA:
			bytesWritten += formula_.Write(data);
			break;
	}
	return bytesWritten;
}
size_t Worksheet::CellTable::RowBlock::CellBlock::DataSize()
{
	switch (type_)
	{
		case CODE::BLANK:
			return blank_.DataSize();
			
		case CODE::BOOLERR:	
			return boolerr_.DataSize();
			
		case CODE::LABELSST:	
			return labelsst_.DataSize();
			
		case CODE::MULBLANK:	
			return mulblank_.DataSize();
			
		case CODE::MULRK:	
			return mulrk_.DataSize();
			
		case CODE::NUMBER:	
			return number_.DataSize();
			
		case CODE::RK:	
			return rk_.DataSize();

		case CODE::FORMULA:
			return formula_.DataSize();
	}
	abort();
}
size_t Worksheet::CellTable::RowBlock::CellBlock::RecordSize() 
{
	switch (type_)
	{
		case CODE::BLANK:
			return blank_.RecordSize();
			
		case CODE::BOOLERR:	
			return boolerr_.RecordSize();
			
		case CODE::LABELSST:	
			return labelsst_.RecordSize();
			
		case CODE::MULBLANK:	
			return mulblank_.RecordSize();
			
		case CODE::MULRK:	
			return mulrk_.RecordSize();
			
		case CODE::NUMBER:	
			return number_.RecordSize();
			
		case CODE::RK:	
			return rk_.RecordSize();

		case CODE::FORMULA:
			return formula_.RecordSize();
	}
	abort();
}
short Worksheet::CellTable::RowBlock::CellBlock::RowIndex()
{
	switch (type_)
	{
		case CODE::BLANK:
			return blank_.rowIndex_;
			
		case CODE::BOOLERR:	
			return boolerr_.rowIndex_;
			
		case CODE::LABELSST:	
			return labelsst_.rowIndex_;
			
		case CODE::MULBLANK:	
			return mulblank_.rowIndex_;
			
		case CODE::MULRK:	
			return mulrk_.rowIndex_;
			
		case CODE::NUMBER:	
			return number_.rowIndex_;
			
		case CODE::RK:	
			return rk_.rowIndex_;

		case CODE::FORMULA:
			return formula_.rowIndex_;
	}
	abort();
}
short Worksheet::CellTable::RowBlock::CellBlock::ColIndex()
{
	switch (type_)
	{
		case CODE::BLANK:
			return blank_.colIndex_;
			
		case CODE::BOOLERR:	
			return boolerr_.colIndex_;
			
		case CODE::LABELSST:	
			return labelsst_.colIndex_;
			
		case CODE::MULBLANK:	
			return mulblank_.firstColIndex_;
			
		case CODE::MULRK:	
			return mulrk_.firstColIndex_;
			
		case CODE::NUMBER:	
			return number_.colIndex_;
			
		case CODE::RK:	
			return rk_.colIndex_;

		case CODE::FORMULA:
			return formula_.colIndex_;
	}
	abort();
}

/************************************************************************************************************/

/************************************************************************************************************/
Worksheet::CellTable::RowBlock::DBCell::DBCell() : Record(),
	firstRowOffset_(0) {code_ = CODE::DBCELL; dataSize_ = 4; recordSize_ = 8;}
size_t Worksheet::CellTable::RowBlock::DBCell::Read(const char* data)
{
	Record::Read(data);
	LittleEndian::Read(data_, firstRowOffset_, 0, 4);
	size_t nm = (dataSize_-4) / 2;
	offsets_.clear();
	offsets_.resize(nm);
	for (size_t i=0; i<nm; ++i)
	{
		LittleEndian::Read(data_, offsets_[i], 4+i*2, 2);
	}
	return RecordSize();
}	
size_t Worksheet::CellTable::RowBlock::DBCell::Write(char* data)
{
	data_.resize(DataSize());
	LittleEndian::Write(data_, firstRowOffset_, 0, 4);
	size_t nm = offsets_.size();
	for (size_t i=0; i<nm; ++i)
	{
		LittleEndian::Write(data_, offsets_[i], 4+i*2, 2);
	}
	return Record::Write(data);
}
size_t Worksheet::CellTable::RowBlock::DBCell::DataSize() {return (dataSize_ = 4+offsets_.size()*2);}
size_t Worksheet::CellTable::RowBlock::DBCell::RecordSize() 
{
	size_t dataSize = DataSize();
	return (recordSize_ = dataSize + 4*(dataSize/8224 + 1));
}
/************************************************************************************************************/

/************************************************************************************************************/
size_t Worksheet::CellTable::RowBlock::Read(const char* data)
{
	size_t bytesRead = 0;
	short code;
	LittleEndian::Read(data, code, 0, 2);
	Row row;
	CellBlock cellBlock;
	cellBlocks_.reserve(1000);
	while (code != CODE::DBCELL)
	{
		switch (code)
		{
			case CODE::ROW:
				rows_.push_back(row);
				bytesRead += rows_.back().Read(data+bytesRead);
				break;
				
			case CODE::BLANK:
			case CODE::BOOLERR:
			case CODE::LABELSST:
			case CODE::MULBLANK:
			case CODE::MULRK:
			case CODE::NUMBER:
			case CODE::RK:
			case CODE::FORMULA:
				cellBlocks_.push_back(cellBlock);
				if (cellBlocks_.size()%1000==0) cellBlocks_.reserve(cellBlocks_.size()+1000);
				bytesRead += cellBlocks_[cellBlocks_.size()-1].Read(data+bytesRead);
				break;

			default:
				Record rec;
				bytesRead += rec.Read(data+bytesRead);
		}	
		LittleEndian::Read(data, code, bytesRead, 2);
	}
	bytesRead += dbcell_.Read(data+bytesRead);
	return bytesRead;
}
size_t Worksheet::CellTable::RowBlock::Write(char* data)
{	
	size_t bytesWritten = 0;
	size_t maxRows = rows_.size();
	{for (size_t i=0; i<maxRows; ++i) 
	{
		bytesWritten += rows_[i].Write(data+bytesWritten);
	}}

	size_t maxCellBlocks = cellBlocks_.size();
	{for (size_t i=0; i<maxCellBlocks; ++i) 
	{
		bytesWritten += cellBlocks_[i].Write(data+bytesWritten);
	}}
	
	bytesWritten += dbcell_.Write(data+bytesWritten);
	return bytesWritten;
}
size_t Worksheet::CellTable::RowBlock::DataSize() 
{
	size_t dataSize = 0;
	size_t maxRows = rows_.size();
	{for (size_t i=0; i<maxRows; ++i) dataSize += rows_[i].RecordSize();}
	
	size_t maxCellBlocks = cellBlocks_.size();
	{for (size_t i=0; i<maxCellBlocks; ++i) dataSize += cellBlocks_[i].RecordSize();}
	
	dataSize += dbcell_.RecordSize();
	return dataSize;	
}
size_t Worksheet::CellTable::RowBlock::RecordSize() {return DataSize();}
/************************************************************************************************************/

/************************************************************************************************************/
size_t Worksheet::CellTable::Read(const char* data)
{
	size_t bytesRead = 0;
	
	short code;
	LittleEndian::Read(data, code, 0, 2);
	RowBlock rowBlock;
	rowBlocks_.reserve(1000);
	while (code == CODE::ROW)
	{
		rowBlocks_.push_back(rowBlock);
		bytesRead += rowBlocks_.back().Read(data+bytesRead);
		LittleEndian::Read(data, code, bytesRead, 2);
	}
	return bytesRead;
}
size_t Worksheet::CellTable::Write(char* data)
{	
	size_t bytesWritten = 0;
	size_t maxRowBlocks_ = rowBlocks_.size();
	for (size_t i=0; i<maxRowBlocks_; ++i) 
	{
		bytesWritten += rowBlocks_[i].Write(data+bytesWritten);
	}
	return bytesWritten;
}
size_t Worksheet::CellTable::DataSize() 
{
	size_t dataSize = 0;
	size_t maxRowBlocks_ = rowBlocks_.size();
	for (size_t i=0; i<maxRowBlocks_; ++i) dataSize += rowBlocks_[i].RecordSize();
	return dataSize;	
}
size_t Worksheet::CellTable::RecordSize() {return DataSize();}
/************************************************************************************************************/

/************************************************************************************************************/
Worksheet::Window2::Window2() : Record(),
	options_(1718), firstVisibleRowIndex_(0), firstVisibleColIndex_(0), gridLineColourIndex_(64),
	unused1_(0), magnificationFactorPageBreakPreview_(0), magnificationFactorNormalView_(0), unused2_(0)
	{code_ = CODE::WINDOW2; dataSize_ = 18; recordSize_ = 22;}

size_t Worksheet::Window2::Read(const char* data)
{
	Record::Read(data);
	LittleEndian::Read(data_, options_, 0, 2);
	LittleEndian::Read(data_, firstVisibleRowIndex_, 2, 2);
	LittleEndian::Read(data_, firstVisibleColIndex_, 4, 2);
	LittleEndian::Read(data_, gridLineColourIndex_, 6, 2);
	LittleEndian::Read(data_, unused1_, 8, 2);
	LittleEndian::Read(data_, magnificationFactorPageBreakPreview_, 10, 2);
	LittleEndian::Read(data_, magnificationFactorNormalView_, 12, 2);
	LittleEndian::Read(data_, unused2_, 14, 4);
	return RecordSize();
}	
size_t Worksheet::Window2::Write(char* data)
{
	data_.resize(DataSize());
	LittleEndian::Write(data_, options_, 0, 2);
	LittleEndian::Write(data_, firstVisibleRowIndex_, 2, 2);
	LittleEndian::Write(data_, firstVisibleColIndex_, 4, 2);
	LittleEndian::Write(data_, gridLineColourIndex_, 6, 2);
	LittleEndian::Write(data_, unused1_, 8, 2);
	LittleEndian::Write(data_, magnificationFactorPageBreakPreview_, 10, 2);
	LittleEndian::Write(data_, magnificationFactorNormalView_, 12, 2);
	LittleEndian::Write(data_, unused2_, 14, 4);
	return Record::Write(data);
}
/************************************************************************************************************/

/************************************************************************************************************/
// Returns true if the supplied rk value contains an integer.
bool IsRKValueAnInteger(int rkValue)
{
	return (rkValue & 2);
}

// Returns true if the supplied rk value contains a double.
bool IsRKValueADouble(int rkValue)
{
	return !(rkValue & 2);
}

// Convert a rk value to a double.
double GetDoubleFromRKValue(int rkValue)
{
	union 
	{
		__int64 intvalue_;
		double doublevalue_;
	} intdouble;

	bool isMultiplied = rkValue & 1;
	rkValue >>= 2;
	intdouble.intvalue_ = rkValue;
	intdouble.intvalue_ <<= 34;
	if (isMultiplied) intdouble.doublevalue_ *= 0.01;
	return intdouble.doublevalue_;
}

// Convert a rk value to an integer.
int GetIntegerFromRKValue(int rkValue)
{
	bool isMultiplied = rkValue & 1;
	rkValue >>= 2;
	if (isMultiplied) rkValue *= 0.01;
	return rkValue;
}

// Convert a double to a rk value.
int GetRKValueFromDouble(double value)
{
	union 
	{
		__int64 intvalue_;
		double doublevalue_;
	} intdouble;

	bool isMultiplied = false;
	int testVal1 = value;
	testVal1 *= 100;
	int testVal2 = value * 100;
	if (testVal1!=testVal2)
	{
		isMultiplied = true;
		value *= 100;
	}

	intdouble.doublevalue_ = value;
	intdouble.intvalue_ >>= 34;

	int rkValue = intdouble.intvalue_;
	rkValue <<= 2;
	rkValue |= isMultiplied;
	return rkValue;
}

// Convert an integer to a rk value.
int GetRKValueFromInteger(int value)
{
	value <<= 2;
	value |= 2;
	return value;
}

// Returns true if the supplied double can be stored as a rk value.
bool CanStoreAsRKValue(double value)
{
	int testVal1 = value * 100;
	testVal1 *= 100;
	int testVal2 = value * 10000;
	if (testVal1!=testVal2) return false;
	else return true;
}

/************************************************************************************************************/

/************************************************************************************************************/
BasicExcel::BasicExcel() {};
BasicExcel::BasicExcel(const char* filename) 
{
	Load(filename);
}

BasicExcel::~BasicExcel() 
{
	if (file_.IsOpen()) file_.Close();
}

// Create a new Excel workbook with a given number of spreadsheets (Minimum 1)
void BasicExcel::New(int sheets)
{
	workbook_ = Workbook();
	worksheets_.clear();

	workbook_.fonts_.resize(4);
	workbook_.XFs_.resize(21);
	workbook_.styles_.resize(6);
	workbook_.boundSheets_.resize(1);
	worksheets_.resize(1);
	UpdateYExcelWorksheet();

	for (int i=0; i<sheets-1; ++i) AddWorksheet();
}

// Load an Excel workbook from a file.
bool BasicExcel::Load(const char* filename)
{
	if (file_.IsOpen()) file_.Close();
	if (file_.Open(filename))
	{
		workbook_ = Workbook();
		worksheets_.clear();

		vector<char> data;
		file_.ReadFile("Workbook", data);
		Read(&*(data.begin()), data.size());
		UpdateYExcelWorksheet();
		return true;
	}
	else return false;
}

// Save current Excel workbook to opened file.
bool BasicExcel::Save()
{
	if (file_.IsOpen())
	{
		// Prepare Raw Worksheets for saving.
		UpdateWorksheets();

		AdjustStreamPositions();	

		// Calculate bytes needed for a workbook.
		size_t minBytes = workbook_.RecordSize();
		size_t maxWorkSheets = worksheets_.size();
		for (size_t i=0; i<maxWorkSheets; ++i)
		{
			minBytes += worksheets_[i].RecordSize();
		}
		
		// Create new workbook.
		vector<char> data(minBytes,0);
		Write(&*(data).begin());

		if (file_.WriteFile("Workbook", data, data.size())!=CompoundFile::SUCCESS) return false;
		return true;
	}
	else return false;
}

// Save current Excel workbook to a file.
bool BasicExcel::SaveAs(const char* filename)
{
	if (file_.IsOpen()) file_.Close();

	if (!file_.Create(filename)) return false;
	if (file_.MakeFile("Workbook")!=CompoundFile::SUCCESS) return false;

	return Save();
}

// Total number of Excel worksheets in current Excel workbook.
size_t BasicExcel::GetTotalWorkSheets()
{
	return worksheets_.size();
}

// Get a pointer to an Excel worksheet at the given index. 
// Index starts from 0. 
// Returns 0 if index is invalid.
BasicExcelWorksheet* BasicExcel::GetWorksheet(size_t sheetIndex)
{
	return &(yesheets_[sheetIndex]);
}

// Get a pointer to an Excel worksheet that has given ANSI name. 
// Returns 0 if there is no Excel worksheet with the given name.
BasicExcelWorksheet* BasicExcel::GetWorksheet(const char* name)
{
	size_t maxWorksheets = yesheets_.size();
	for (size_t i=0; i<maxWorksheets; ++i)
	{
		if (workbook_.boundSheets_[i].name_.unicode_ & 1) continue;
		if (strcmp(name, workbook_.boundSheets_[i].name_.name_) == 0) return &(yesheets_[i]);
	}
	return 0;
}

// Get a pointer to an Excel worksheet that has given Unicode name. 
// Returns 0 if there is no Excel worksheet with the given name.
BasicExcelWorksheet* BasicExcel::GetWorksheet(const wchar_t* name)
{
	size_t maxWorksheets = yesheets_.size();
	for (size_t i=0; i<maxWorksheets; ++i)
	{
		if (!(workbook_.boundSheets_[i].name_.unicode_ & 1)) continue;
		if (wcscmp(name, workbook_.boundSheets_[i].name_.wname_) == 0) return &(yesheets_[i]);
	}
	return 0;
}

// Add a new Excel worksheet to the given index. 
// Name given to worksheet is SheetX, where X is a number which starts from 1. 
// Index starts from 0. 
// Worksheet is added to the last position if sheetIndex == -1. 
// Returns a pointer to the worksheet if successful, 0 if otherwise.
BasicExcelWorksheet* BasicExcel::AddWorksheet(int sheetIndex)
{
	size_t sheetNo = yesheets_.size() + 1;
	BasicExcelWorksheet* yesheet = 0;
	do
	{
		char sname[50];
		sprintf(sname, "Sheet%d", sheetNo++);
		yesheet = AddWorksheet(sname, sheetIndex);
	} while (!yesheet);
	return yesheet;
}

// Add a new Excel worksheet with given ANSI name to the given index.
// Index starts from 0. 
// Worksheet is added to the last position if sheetIndex == -1. 
// Returns a pointer to the worksheet if successful, 0 if otherwise.
BasicExcelWorksheet* BasicExcel::AddWorksheet(const char* name, int sheetIndex)
{
	size_t maxWorksheets = yesheets_.size();
	for (size_t i=0; i<maxWorksheets; ++i)
	{
		if (workbook_.boundSheets_[i].name_.unicode_ & 1) continue;
		if (strcmp(name, workbook_.boundSheets_[i].name_.name_) == 0) return 0;
	}

	Workbook::BoundSheet* boundSheet;
	Worksheet* worksheet;
	BasicExcelWorksheet* yesheet;
	if (sheetIndex == -1)
	{
		workbook_.boundSheets_.push_back(Workbook::BoundSheet());
		worksheets_.push_back(Worksheet());
		yesheets_.push_back(BasicExcelWorksheet(this, worksheets_.size()-1));
		boundSheet = &(workbook_.boundSheets_.back());
		worksheet = &(worksheets_.back());
		yesheet = &(yesheets_.back());
	}
	else
	{
		boundSheet = &*(workbook_.boundSheets_.insert(workbook_.boundSheets_.begin()+sheetIndex, Workbook::BoundSheet()));
		worksheet = &*(worksheets_.insert(worksheets_.begin()+sheetIndex, Worksheet()));
		yesheet = &*(yesheets_.insert(yesheets_.begin()+sheetIndex, BasicExcelWorksheet(this, sheetIndex)));
		size_t maxSheets = worksheets_.size();
		for (size_t i=sheetIndex+1; i<maxSheets; ++i)
		{
			yesheets_[i].sheetIndex_ = i;
		}
	}
	boundSheet->name_ = name;
	worksheet->window2_.options_ &= ~0x200;
	return yesheet;
}

// Add a new Excel worksheet with given Unicode name to the given index.
// Index starts from 0.
// Worksheet is added to the last position if sheetIndex == -1.
// Returns a pointer to the worksheet if successful, 0 if otherwise.
BasicExcelWorksheet* BasicExcel::AddWorksheet(const wchar_t* name, int sheetIndex)
{
	size_t maxWorksheets = yesheets_.size();
	for (size_t i=0; i<maxWorksheets; ++i)
	{
		if (!(workbook_.boundSheets_[i].name_.unicode_ & 1)) continue;
		if (wcscmp(name, workbook_.boundSheets_[i].name_.wname_) == 0) return 0;
	}

	Workbook::BoundSheet* boundSheet;
	Worksheet* worksheet;
	BasicExcelWorksheet* yesheet;
	if (sheetIndex == -1)
	{
		workbook_.boundSheets_.push_back(Workbook::BoundSheet());
		worksheets_.push_back(Worksheet());
		yesheets_.push_back(BasicExcelWorksheet(this, worksheets_.size()-1));
		boundSheet = &(workbook_.boundSheets_.back());
		worksheet = &(worksheets_.back());
		yesheet = &(yesheets_.back());
	}
	else
	{
		boundSheet = &*(workbook_.boundSheets_.insert(workbook_.boundSheets_.begin()+sheetIndex, Workbook::BoundSheet()));
		worksheet = &*(worksheets_.insert(worksheets_.begin()+sheetIndex, Worksheet()));
		yesheet = &*(yesheets_.insert(yesheets_.begin()+sheetIndex, BasicExcelWorksheet(this, sheetIndex)));
		size_t maxSheets = worksheets_.size();
		for (size_t i=sheetIndex+1; i<maxSheets; ++i)
		{
			yesheets_[i].sheetIndex_ = i;
		}
	}
	boundSheet->name_ = name;
	worksheet->window2_.options_ &= ~0x200;
	return yesheet;
}

// Delete an Excel worksheet at the given index.
// Index starts from 0.
// Returns true if successful, false if otherwise.
bool BasicExcel::DeleteWorksheet(size_t sheetIndex)
{
	if (sheetIndex<workbook_.boundSheets_.size())
	{
		workbook_.boundSheets_.erase(workbook_.boundSheets_.begin()+sheetIndex);
		worksheets_.erase(worksheets_.begin()+sheetIndex);
		yesheets_.erase(yesheets_.begin()+sheetIndex);
		return true;
	}
	else return false;
}

// Delete an Excel worksheet that has given ANSI name.
// Returns true if successful, false if otherwise.
bool BasicExcel::DeleteWorksheet(const char* name)
{
	size_t maxWorksheets = yesheets_.size();
	for (size_t i=0; i<maxWorksheets; ++i)
	{
		if (workbook_.boundSheets_[i].name_.unicode_ & 1) continue;
		if (strcmp(name, workbook_.boundSheets_[i].name_.name_) == 0) return DeleteWorksheet(i);
	}
	return false;
}

// Delete an Excel worksheet that has given Unicode name.
// Returns true if successful, false if otherwise.
bool BasicExcel::DeleteWorksheet(const wchar_t* name)
{
	size_t maxWorksheets = worksheets_.size();
	for (size_t i=0; i<maxWorksheets; ++i)
	{
		if (!(workbook_.boundSheets_[i].name_.unicode_ & 1)) continue;
		if (wcscmp(name, workbook_.boundSheets_[i].name_.wname_) == 0) return DeleteWorksheet(i);
	}
	return false;
}

// Get the worksheet name at the given index.
// Index starts from 0.
// Returns 0 if name is in Unicode format.
char* BasicExcel::GetAnsiSheetName(size_t sheetIndex)
{
	if (!(workbook_.boundSheets_[sheetIndex].name_.unicode_ & 1))
	{
		return workbook_.boundSheets_[sheetIndex].name_.name_;
	}
	else return 0;
}

// Get the worksheet name at the given index.
// Index starts from 0.
// Returns 0 if name is in Ansi format.
wchar_t* BasicExcel::GetUnicodeSheetName(size_t sheetIndex)
{
	if (workbook_.boundSheets_[sheetIndex].name_.unicode_ & 1)
	{
		return workbook_.boundSheets_[sheetIndex].name_.wname_;
	}
	else return 0;
}

// Get the worksheet name at the given index.
// Index starts from 0.
// Returns false if name is in Unicode format.
bool BasicExcel::GetSheetName(size_t sheetIndex, char* name)
{
	if (!(workbook_.boundSheets_[sheetIndex].name_.unicode_ & 1))
	{
		strcpy(name, workbook_.boundSheets_[sheetIndex].name_.name_);
		return true;
	}
	else return false;
}

// Get the worksheet name at the given index.
// Index starts from 0.
// Returns false if name is in Ansi format.
bool BasicExcel::GetSheetName(size_t sheetIndex, wchar_t* name)
{
	if (workbook_.boundSheets_[sheetIndex].name_.unicode_ & 1)
	{
		wcscpy(name, workbook_.boundSheets_[sheetIndex].name_.wname_);
		return true;
	}
	else return false;
}

// Rename an Excel worksheet at the given index to the given ANSI name.
// Index starts from 0.
// Returns true if successful, false if otherwise.
bool BasicExcel::RenameWorksheet(size_t sheetIndex, const char* to)
{
	size_t maxWorksheets = yesheets_.size();
	if (sheetIndex < maxWorksheets)
	{
		for (size_t i=0; i<maxWorksheets; ++i)
		{
			if (workbook_.boundSheets_[i].name_.unicode_ & 1) continue;
			if (strcmp(to, workbook_.boundSheets_[i].name_.name_) == 0) return false;
		}
		workbook_.boundSheets_[sheetIndex].name_ = to;
		return true;
	}
	else return false;
}

// Rename an Excel worksheet at the given index to the given Unicode name.
// Index starts from 0.
// Returns true if successful, false if otherwise.
bool BasicExcel::RenameWorksheet(size_t sheetIndex, const wchar_t* to)
{
	size_t maxWorksheets = yesheets_.size();
	if (sheetIndex < maxWorksheets)
	{
		for (size_t i=0; i<maxWorksheets; ++i)
		{
			if (!(workbook_.boundSheets_[i].name_.unicode_ & 1)) continue;
			if (wcscmp(to, workbook_.boundSheets_[i].name_.wname_) == 0) return false;
		}
		workbook_.boundSheets_[sheetIndex].name_ = to;
		return true;
	}
	else return false;
}

// Rename an Excel worksheet that has given ANSI name to another ANSI name.
// Returns true if successful, false if otherwise.
bool BasicExcel::RenameWorksheet(const char* from, const char* to)
{
	size_t maxWorksheets = yesheets_.size();
	for (size_t i=0; i<maxWorksheets; ++i)
	{
		if (workbook_.boundSheets_[i].name_.unicode_ & 1) continue;
		if (strcmp(from, workbook_.boundSheets_[i].name_.name_) == 0)
		{
			for (size_t j=0; j<maxWorksheets; ++j)
			{
				if (workbook_.boundSheets_[j].name_.unicode_ & 1) continue;
				if (strcmp(to, workbook_.boundSheets_[j].name_.name_) == 0) return false;
			}
			workbook_.boundSheets_[i].name_ = to;
			return true;
		}
	}
	return false;
}

// Rename an Excel worksheet that has given Unicode name to another Unicode name.
// Returns true if successful, false if otherwise.
bool BasicExcel::RenameWorksheet(const wchar_t* from, const wchar_t* to)
{
	size_t maxWorksheets = worksheets_.size();
	for (size_t i=0; i<maxWorksheets; ++i)
	{
		if (!(workbook_.boundSheets_[i].name_.unicode_ & 1)) continue;
		if (wcscmp(from, workbook_.boundSheets_[i].name_.wname_) == 0)
		{
			for (size_t j=0; j<maxWorksheets; ++j)
			{
				if (!(workbook_.boundSheets_[j].name_.unicode_ & 1)) continue;
				if (wcscmp(to, workbook_.boundSheets_[j].name_.wname_) == 0) return false;
			}
			workbook_.boundSheets_[i].name_ = to;
			return true;
		}
	}
	return false;
}

size_t BasicExcel::Read(const char* data, size_t dataSize)
{
	size_t bytesRead = 0;
	BOF bof;
	short code;
	LittleEndian::Read(data, code, 0, 2);
	Record rec;
	while (code == CODE::BOF)
	{
		bof.Read(data+bytesRead);
		switch (bof.type_)
		{
			case WORKBOOK_GLOBALS:
				bytesRead += workbook_.Read(data+bytesRead);
				break;
				
			case VISUAL_BASIC_MODULE:
				bytesRead += rec.Read(data+bytesRead);
				break;
				
			case WORKSHEET:
				worksheets_.push_back(Worksheet());
				bytesRead += worksheets_.back().Read(data+bytesRead);
				break;
				
			case CHART:
				bytesRead += rec.Read(data+bytesRead);
				break;		

			default:
				bytesRead += rec.Read(data+bytesRead);
				break;
		}
		if (bytesRead < dataSize) LittleEndian::Read(data, code, bytesRead, 2);
		else break;
	} 
	return bytesRead;
}

size_t BasicExcel::Write(char* data)
{
	size_t bytesWritten = 0;
	bytesWritten += workbook_.Write(data+bytesWritten);
	
	size_t maxWorkSheets = worksheets_.size();
	for (size_t i=0; i<maxWorkSheets; ++i)
	{
		bytesWritten += worksheets_[i].Write(data+bytesWritten);
	}
	return bytesWritten;
}

void BasicExcel::AdjustStreamPositions()
{
//	AdjustExtSSTPositions();
	AdjustBoundSheetBOFPositions();
	AdjustDBCellPositions();
}

void BasicExcel::AdjustBoundSheetBOFPositions()
{
	size_t offset = workbook_.RecordSize();
	size_t maxBoundSheets = workbook_.boundSheets_.size();
	for (size_t i=0; i<maxBoundSheets; ++i)
	{
		workbook_.boundSheets_[i].BOFpos_ = offset;
		offset += worksheets_[i].RecordSize();
	}
}

void BasicExcel::AdjustDBCellPositions()
{
	size_t offset = workbook_.RecordSize();
	size_t maxSheets = worksheets_.size();
	for (size_t i=0; i<maxSheets; ++i)
	{
		offset += worksheets_[i].bof_.RecordSize();
		offset += worksheets_[i].index_.RecordSize();
		offset += worksheets_[i].dimensions_.RecordSize();
		
		size_t maxRowBlocks_ = worksheets_[i].cellTable_.rowBlocks_.size();
		for (size_t j=0; j<maxRowBlocks_; ++j) 
		{
			size_t firstRowOffset = 0;

			size_t maxRows = worksheets_[i].cellTable_.rowBlocks_[j].rows_.size();
			{for (size_t k=0; k<maxRows; ++k) 
			{
				offset += worksheets_[i].cellTable_.rowBlocks_[j].rows_[k].RecordSize();
				firstRowOffset += worksheets_[i].cellTable_.rowBlocks_[j].rows_[k].RecordSize();
			}}
			size_t cellOffset = firstRowOffset - 20; // a ROW record is 20 bytes long

			size_t maxCellBlocks = worksheets_[i].cellTable_.rowBlocks_[j].cellBlocks_.size();
			{for (size_t k=0; k<maxCellBlocks; ++k) 
			{
				offset += worksheets_[i].cellTable_.rowBlocks_[j].cellBlocks_[k].RecordSize();
				firstRowOffset += worksheets_[i].cellTable_.rowBlocks_[j].cellBlocks_[k].RecordSize();
			}}

			// Adjust Index DBCellPos_ absolute offset
			worksheets_[i].index_.DBCellPos_[j] = offset; 
			
			offset += worksheets_[i].cellTable_.rowBlocks_[j].dbcell_.RecordSize();

			// Adjust DBCell first row offsets
			worksheets_[i].cellTable_.rowBlocks_[j].dbcell_.firstRowOffset_ = firstRowOffset;

			// Adjust DBCell offsets
			size_t l=0;
			{for (size_t k=0; k<maxRows; ++k)
			{
				for (; l<maxCellBlocks; ++l)
				{
					if (worksheets_[i].cellTable_.rowBlocks_[j].rows_[k].rowIndex_ <=
						worksheets_[i].cellTable_.rowBlocks_[j].cellBlocks_[l].RowIndex())
					{
						worksheets_[i].cellTable_.rowBlocks_[j].dbcell_.offsets_[k] = cellOffset;
						break;
					}
					cellOffset += worksheets_[i].cellTable_.rowBlocks_[j].cellBlocks_[l].RecordSize();
				}
				cellOffset = 0;
			}}
		}	
		
		offset += worksheets_[i].cellTable_.RecordSize();
		offset += worksheets_[i].window2_.RecordSize();
		offset += worksheets_[i].eof_.RecordSize();
	}
}

void BasicExcel::AdjustExtSSTPositions()
{
	size_t offset = workbook_.bof_.RecordSize();
	offset += workbook_.bof_.RecordSize();
	offset += workbook_.window1_.RecordSize();

	size_t maxFonts = workbook_.fonts_.size();
	{for (size_t i=0; i<maxFonts; ++i) {offset += workbook_.fonts_[i].RecordSize();}}
	
	size_t maxXFs = workbook_.XFs_.size();
	{for (size_t i=0; i<maxXFs; ++i) {offset += workbook_.XFs_[i].RecordSize();}}

	size_t maxStyles = workbook_.styles_.size();
	{for (size_t i=0; i<maxStyles; ++i) {offset += workbook_.styles_[i].RecordSize();}}

	size_t maxBoundSheets = workbook_.boundSheets_.size();
	{for (size_t i=0; i<maxBoundSheets; ++i) {offset += workbook_.boundSheets_[i].RecordSize();}}

	workbook_.extSST_.stringsTotal_ = 10; 
	size_t maxPortions = workbook_.sst_.uniqueStringsTotal_ / workbook_.extSST_.stringsTotal_ +
						(workbook_.sst_.uniqueStringsTotal_%workbook_.extSST_.stringsTotal_ ? 1 : 0);
	workbook_.extSST_.streamPos_.resize(maxPortions);
	workbook_.extSST_.firstStringPos_.resize(maxPortions);
	workbook_.extSST_.unused_.resize(maxPortions);

	size_t relativeOffset = 8;
	for (size_t i=0; i<maxPortions; ++i)
	{
		workbook_.extSST_.streamPos_[i] = offset + 4 + relativeOffset;
		workbook_.extSST_.firstStringPos_[i] = 4 + relativeOffset;
		workbook_.extSST_.unused_[i] = 0;		

		for (size_t j=0; j<workbook_.extSST_.stringsTotal_; ++j)
		{
			if (i*workbook_.extSST_.stringsTotal_+j >= workbook_.sst_.strings_.size()) break;
			size_t stringSize = workbook_.sst_.strings_[i*workbook_.extSST_.stringsTotal_+j].StringSize();
			if (relativeOffset+stringSize+3 < 8224)
			{
				relativeOffset += stringSize + 3;
			}
			else
			{
				// If have >= 12 bytes (2 for size, 1 for unicode and >=9 for data, can split string
				// otherwise, end record and start continue record.
				if (8224 - relativeOffset >= 12)
				{
					stringSize -= (8224 - relativeOffset - 3);
					offset += 12 + relativeOffset;
					relativeOffset = 0;

					size_t additionalContinueRecords = stringSize / 8223; // 8223 because the first byte is for unicode
					for (size_t k=0; k<additionalContinueRecords; ++k)
					{
						stringSize -= 8223;
					}
					relativeOffset += stringSize + 1;
				}
				else
				{
					if (relativeOffset+stringSize+3 < 8224)
					{
						relativeOffset += stringSize + 3;
					}
					else
					{
						// If have >= 12 bytes (2 for size, 1 for unicode and >=9 for data, can split string
						// otherwise, end record and start continue record.
						if (8224 - relativeOffset >= 12)
						{
							stringSize -= (8224 - relativeOffset - 3);
							offset += 12 + relativeOffset;
							relativeOffset = 0;

							size_t additionalContinueRecords = stringSize / 8223; // 8223 because the first byte is for unicode
							for (size_t k=0; k<additionalContinueRecords; ++k)
							{
								stringSize -= 8223;
							}
							relativeOffset += stringSize + 1;
						}
					}
				}			
			}		
		}
	}
}

// Update yesheets_ using information from worksheets_.
void BasicExcel::UpdateYExcelWorksheet()
{
	size_t maxWorksheets = worksheets_.size();
	yesheets_.clear();
	yesheets_.reserve(maxWorksheets);
	for (size_t i=0; i<maxWorksheets; ++i)
	{
		yesheets_.push_back(BasicExcelWorksheet(this,i));
	}
}

// Update worksheets_ using information from yesheets_.
void BasicExcel::UpdateWorksheets()
{
	// Constants.
	const size_t maxWorksheets = yesheets_.size();
	Worksheet::CellTable::RowBlock rowBlock;
	Worksheet::CellTable::RowBlock::CellBlock cellBlock;
	Worksheet::CellTable::RowBlock::Row row;
	Worksheet::CellTable::RowBlock::CellBlock::MulRK::XFRK xfrk;
	LargeString largeString;

	map<vector<char>, size_t> stringMap;
	map<vector<char>, size_t>::iterator stringMapIt;
	map<vector<wchar_t>, size_t> wstringMap;
	map<vector<wchar_t>, size_t>::iterator wstringMapIt;

	// Reset worksheets and string table.
	worksheets_.clear();
	worksheets_.resize(maxWorksheets);
	
	workbook_.sst_.stringsTotal_ = 0;
	workbook_.sst_.uniqueStringsTotal_ = 0;
	workbook_.sst_.strings_.clear();

	for (size_t s=0; s<maxWorksheets; ++s)
	{
		size_t maxRows = yesheets_[s].GetTotalRows();
		size_t maxCols = yesheets_[s].GetTotalCols();

		// Modify Index
		worksheets_[s].index_.firstUsedRowIndex_ = 100000;	// Use 100000 to indicate that firstUsedRowIndex is not set yet since maximum allowed rows in Excel is 65535.
		worksheets_[s].index_.firstUnusedRowIndex_ = maxRows;

		// Modify Dimensions
		worksheets_[s].dimensions_.firstUsedRowIndex_ = 100000; // Use 100000 to indicate that firstUsedRowIndex is not set yet since maximum allowed rows in Excel is 65535.
		worksheets_[s].dimensions_.firstUsedColIndex_ = 1000;	// Use 1000 to indicate that firstUsedColIndex is not set yet since maximum allowed columns in Excel is 255.
		worksheets_[s].dimensions_.lastUsedRowIndexPlusOne_ = maxRows;
		worksheets_[s].dimensions_.lastUsedColIndexPlusOne_ = maxCols;

		// Make first sheet selected and other sheets unselected
		if (s > 0) worksheets_[s].window2_.options_ &= ~0x200;

		// References and pointers to shorten code
		vector<Worksheet::CellTable::RowBlock>& rRowBlocks = worksheets_[s].cellTable_.rowBlocks_;
		vector<Worksheet::CellTable::RowBlock::CellBlock>* pCellBlocks;
		Worksheet::CellTable::RowBlock::CellBlock* pCell;
		rRowBlocks.resize(maxRows/32 + (maxRows%32 ? 1 : 0));
		for (size_t r=0, curRowBlock=0; r<maxRows; ++r)
		{
			if (r%32==0) 
			{
				// New row block for every 32 rows.
				pCellBlocks = &(rRowBlocks[curRowBlock++].cellBlocks_);								
			}
			bool newRow = true;	// Keep track whether current row contains data.
			pCellBlocks->reserve(1000);
			for (size_t c=0; c<maxCols; ++c)
			{
				BasicExcelCell* cell = yesheets_[s].Cell(r,c);
				int cellType = cell->Type();
				if (cellType != BasicExcelCell::UNDEFINED)	// Current cell contains some data
				{		
					if (worksheets_[s].index_.firstUsedRowIndex_ == 100000) 
					{
						// Set firstUsedRowIndex.
						worksheets_[s].index_.firstUsedRowIndex_ = r; 
						worksheets_[s].dimensions_.firstUsedRowIndex_ = r;

						// Resize DBCellPos.
						size_t nm = int(worksheets_[s].index_.firstUnusedRowIndex_ - worksheets_[s].index_.firstUsedRowIndex_ - 1) / 32 + 1;
						worksheets_[s].index_.DBCellPos_.resize(nm);
					}
					if (worksheets_[s].dimensions_.firstUsedColIndex_ == 1000)
					{
						// Set firstUsedColIndex.
						worksheets_[s].dimensions_.firstUsedColIndex_ = c;
					}

					if (newRow)
					{
						// Prepare Row and DBCell for new row with data.
						Worksheet::CellTable::RowBlock& rRowBlock = rRowBlocks[curRowBlock-1];
						rRowBlock.rows_.push_back(row);
						rRowBlock.rows_.back().rowIndex_ = r;
						rRowBlock.rows_.back().lastCellColIndexPlusOne_ = maxCols;
						rRowBlock.dbcell_.offsets_.push_back(0);
						newRow = false;
					}

					// Create new cellblock to store cell.
					pCellBlocks->push_back(cellBlock);
					if (pCellBlocks->size()%1000==0) pCellBlocks->reserve(pCellBlocks->size()+1000);
					pCell = &(pCellBlocks->back());

					// Store cell.
					switch(cellType)
					{
						case BasicExcelCell::INT:
						{
							// Check whether it is a single cell or range of cells.
							size_t cl = c + 1;
							for (; cl<maxCols; ++cl)
							{
								BasicExcelCell* cellNext = yesheets_[s].Cell(r,cl);
								if (cellNext->Type()==BasicExcelCell::UNDEFINED ||
									cellNext->Type()!=cell->Type()) break;
							}

							if (cl > c+1)
							{
								// MULRK cells
								pCell->type_ = CODE::MULRK;
								pCell->normalType_ = true;
								pCell->mulrk_.rowIndex_ = r;
								pCell->mulrk_.firstColIndex_ = c;
								pCell->mulrk_.lastColIndex_ = cl - 1;
								pCell->mulrk_.XFRK_.resize(cl-c);
								for (size_t i=0; c<cl; ++c, ++i)
								{
									cell = yesheets_[s].Cell(r,c);
									pCell->mulrk_.XFRK_[i].RKValue_ = GetRKValueFromInteger(cell->GetInteger());
								}
								--c;
							}
							else
							{
								// Single cell
								pCell->normalType_ = true;
								pCell->type_ = CODE::RK;
								pCell->rk_.rowIndex_ = r;
								pCell->rk_.colIndex_ = c;
								pCell->rk_.value_ = GetRKValueFromInteger(cell->GetInteger());
							}
							break;
						}

						case BasicExcelCell::DOUBLE:
						{
							// Check whether it is a single cell or range of cells.
							// Double values which cannot be stored as RK values will be stored as single cells.
							bool canStoreAsRKValue = CanStoreAsRKValue(cell->GetDouble());
							size_t cl = c + 1;
							for (; cl<maxCols; ++cl)
							{
								BasicExcelCell* cellNext = yesheets_[s].Cell(r,cl);
								if (cellNext->Type()==BasicExcelCell::UNDEFINED ||
									cellNext->Type()!=cell->Type() ||
									canStoreAsRKValue!=CanStoreAsRKValue(cellNext->GetDouble())) break;
							}

							if (cl > c+1 && canStoreAsRKValue)
							{
								// MULRK cells
								pCell->type_ = CODE::MULRK;
								pCell->normalType_ = true;
								pCell->mulrk_.rowIndex_ = r;
								pCell->mulrk_.firstColIndex_ = c;
								pCell->mulrk_.lastColIndex_ = cl - 1;
								pCell->mulrk_.XFRK_.resize(cl-c);
								for (size_t i=0; c<cl; ++c, ++i)
								{
									cell = yesheets_[s].Cell(r,c);
									pCell->mulrk_.XFRK_[i].RKValue_ = GetRKValueFromDouble(cell->GetDouble());
								}
								--c;
							}
							else
							{
								// Single cell
								pCell->normalType_ = true;
								if (canStoreAsRKValue)
								{
									pCell->type_ = CODE::RK;
									pCell->rk_.rowIndex_ = r;
									pCell->rk_.colIndex_ = c;
									pCell->rk_.value_ = GetRKValueFromDouble(cell->GetDouble());
								}
								else
								{									
									pCell->type_ = CODE::NUMBER;
									pCell->number_.rowIndex_ = r;
									pCell->number_.colIndex_ = c;
									pCell->number_.value_ = cell->GetDouble();								
								}
							}
							break;
						}

						case BasicExcelCell::STRING:
						{
							// Fill cell information
							pCell->type_ = CODE::LABELSST;
							pCell->normalType_ = true;
							pCell->labelsst_.rowIndex_ = r;
							pCell->labelsst_.colIndex_ = c;
							
							// Get cell string
							vector<char> str(cell->GetStringLength()+1);
							cell->Get(&*(str.begin()));
							str.pop_back();	// Remove null character because LargeString does not store null character.

							// Check if string is present in Shared string table.
							++workbook_.sst_.stringsTotal_;
							size_t maxUniqueStrings = workbook_.sst_.uniqueStringsTotal_;
							size_t strIndex = 0;
							stringMapIt = stringMap.find(str);
							if (stringMapIt != stringMap.end()) strIndex = stringMapIt->second;
							else strIndex = maxUniqueStrings;

							if (strIndex < maxUniqueStrings)
							{
								// String is present in Shared string table.
								pCell->labelsst_.SSTRecordIndex_ = strIndex;
							}
							else
							{
								// New unique string.
								stringMap[str] = maxUniqueStrings;
								workbook_.sst_.strings_.push_back(largeString);
								workbook_.sst_.strings_[maxUniqueStrings].name_ = str;
								workbook_.sst_.strings_[maxUniqueStrings].unicode_ = 0;
								pCell->labelsst_.SSTRecordIndex_ = maxUniqueStrings;
								++workbook_.sst_.uniqueStringsTotal_;
							}
							break;
						}

						case BasicExcelCell::WSTRING:
						{
							// Fill cell information
							pCell->type_ = CODE::LABELSST;
							pCell->normalType_ = true;
							pCell->labelsst_.rowIndex_ = r;
							pCell->labelsst_.colIndex_ = c;

							// Get cell string
							vector<wchar_t> str(cell->GetStringLength()+1);
							cell->Get(&*(str.begin()));
							str.pop_back(); // Remove null character because LargeString does not store null character.

							// Check if string is present in Shared string table.
							++workbook_.sst_.stringsTotal_;
							size_t maxUniqueStrings = workbook_.sst_.strings_.size();
							size_t strIndex = 0;
							wstringMapIt = wstringMap.find(str);
							if (wstringMapIt != wstringMap.end()) strIndex = wstringMapIt->second;
							else strIndex = maxUniqueStrings;

							if (strIndex < maxUniqueStrings)
							{
								// String is present in Shared string table.
								pCell->labelsst_.SSTRecordIndex_ = strIndex;
							}
							else
							{
								// New unique string
								wstringMap[str] = maxUniqueStrings;
								workbook_.sst_.strings_.push_back(largeString);
								workbook_.sst_.strings_[maxUniqueStrings].wname_ = str;
								workbook_.sst_.strings_[maxUniqueStrings].unicode_ = 1;
								pCell->labelsst_.SSTRecordIndex_ = maxUniqueStrings;
								++workbook_.sst_.uniqueStringsTotal_;
							}
							break;
						}
					}
				}
			}
		}

		// If worksheet has no data
		if (worksheets_[s].index_.firstUsedRowIndex_ == 100000) 
		{
			// Set firstUsedRowIndex.
			worksheets_[s].index_.firstUsedRowIndex_ = 0; 
			worksheets_[s].dimensions_.firstUsedRowIndex_ = 0;

			// Resize DBCellPos.
			size_t nm = int(worksheets_[s].index_.firstUnusedRowIndex_ - worksheets_[s].index_.firstUsedRowIndex_ - 1) / 32 + 1;
			worksheets_[s].index_.DBCellPos_.resize(nm);
		}
		if (worksheets_[s].dimensions_.firstUsedColIndex_ == 1000)
		{
			// Set firstUsedColIndex.
			worksheets_[s].dimensions_.firstUsedColIndex_ = 0;
		}
	}
}
/************************************************************************************************************/

/************************************************************************************************************/
BasicExcelWorksheet::BasicExcelWorksheet(BasicExcel* excel, size_t sheetIndex) : 
	excel_(excel), sheetIndex_(sheetIndex) 
{
	UpdateCells();
}

// Get the current worksheet name.
// Returns 0 if name is in Unicode format.
char* BasicExcelWorksheet::GetAnsiSheetName()
{
	if (!(excel_->workbook_.boundSheets_[sheetIndex_].name_.unicode_ & 1))
	{
		return excel_->workbook_.boundSheets_[sheetIndex_].name_.name_;
	}
	else return 0;
}

// Get the current worksheet name.
// Returns 0 if name is in Ansi format.
wchar_t* BasicExcelWorksheet::GetUnicodeSheetName()
{
	if (excel_->workbook_.boundSheets_[sheetIndex_].name_.unicode_ & 1)
	{
		return excel_->workbook_.boundSheets_[sheetIndex_].name_.wname_;
	}
	else return 0;
}

// Get the current worksheet name.
// Returns false if name is in Unicode format.
bool BasicExcelWorksheet::GetSheetName(char* name)
{
	if (!(excel_->workbook_.boundSheets_[sheetIndex_].name_.unicode_ & 1))
	{
		strcpy(name, excel_->workbook_.boundSheets_[sheetIndex_].name_.name_);
		return true;
	}
	else return false;
}

// Get the current worksheet name.
// Returns false if name is in Ansi format.
bool BasicExcelWorksheet::GetSheetName(wchar_t* name)
{
	if (excel_->workbook_.boundSheets_[sheetIndex_].name_.unicode_ & 1)
	{
		wcscpy(name, excel_->workbook_.boundSheets_[sheetIndex_].name_.wname_);
		return true;
	}
	else return false;
}

// Rename current Excel worksheet to another ANSI name.
// Returns true if successful, false if otherwise.
bool BasicExcelWorksheet::Rename(const char* to)
{
	size_t maxWorksheets = excel_->workbook_.boundSheets_.size();
	for (size_t i=0; i<maxWorksheets; ++i)
	{
		if (excel_->workbook_.boundSheets_[i].name_.unicode_ & 1) continue;
		if (strcmp(to, excel_->workbook_.boundSheets_[i].name_.name_) == 0) return false;
	}

	excel_->workbook_.boundSheets_[sheetIndex_].name_ = to;
	return true;
}

// Rename current Excel worksheet to another Unicode name.
// Returns true if successful, false if otherwise.
bool BasicExcelWorksheet::Rename(const wchar_t* to)
{
	size_t maxWorksheets = excel_->workbook_.boundSheets_.size();
	for (size_t i=0; i<maxWorksheets; ++i)
	{
		if (!(excel_->workbook_.boundSheets_[i].name_.unicode_ & 1)) continue;
		if (wcscmp(to, excel_->workbook_.boundSheets_[i].name_.wname_) == 0) return false;
	}

	excel_->workbook_.boundSheets_[sheetIndex_].name_ = to;
	return true;
}

///< Print entire worksheet to an output stream, separating each column with the defined delimiter and enclosing text using the defined textQualifier.
///< Leave out the textQualifier argument if do not wish to have any text qualifiers.
void BasicExcelWorksheet::Print(ostream& os, char delimiter, char textQualifier)
{
	for (size_t r=0; r<maxRows_; ++r)
	{
		for (size_t c=0; c<maxCols_; ++c)
		{
			BasicExcelCell* cell = Cell(r,c);
			switch (cell->Type())
			{
				case BasicExcelCell::UNDEFINED:
					break;

				case BasicExcelCell::INT:
					os << cell->GetInteger();
					break;

				case BasicExcelCell::DOUBLE:
					os << setprecision(15) << cell->GetDouble();
					break;

				case BasicExcelCell::STRING:
				{
					if (textQualifier != '\0')
						{
						// Get string.
						size_t maxLength = cell->GetStringLength();
						vector<char> cellString(maxLength+1);
						cell->Get(&*(cellString.begin()));

						// Duplicate textQualifier if found in string.
						vector<char>::iterator it;
						size_t npos = 0;
						while ((it=find(cellString.begin()+npos, cellString.end(), textQualifier)) != cellString.end())
						{
							npos = distance(cellString.begin(), cellString.insert(it, textQualifier)) + 2;
						}

						// Print out string enclosed with textQualifier.
						os << textQualifier << &*(cellString.begin()) << textQualifier;
					}
					else os << cell->GetString();
					break;
				}

				case BasicExcelCell::WSTRING:
				{
					// Print out string enclosed with textQualifier (does not work).
					//os << textQualifier << cell->GetWString() << textQualifier;
					break;
				}
			}
			if (c < maxCols_-1) os << delimiter;
		}
		os << endl;
	}
}

// Total number of rows in current Excel worksheet.
size_t BasicExcelWorksheet::GetTotalRows()
{
	return maxRows_;
}

// Total number of columns in current Excel worksheet.
size_t BasicExcelWorksheet::GetTotalCols()
{
	return maxCols_;
}

// Return a pointer to an Excel cell.
// row and col starts from 0.
// Returns 0 if row exceeds 65535 or col exceeds 255.
BasicExcelCell* BasicExcelWorksheet::Cell(size_t row, size_t col)	
{
	// Check to ensure row and col does not exceed maximum allowable range for an Excel worksheet.
	if (row>65535 || col>255) return 0;

	// Increase size of cells matrix if necessary
	if (col>=maxCols_)
	{
		// Increase number of columns.
		maxCols_ = col + 1;
		for (size_t i=0; i<maxRows_; ++i) cells_[i].resize(maxCols_);
	}
	if (row>=maxRows_) 
	{
		// Increase number of rows.
		maxRows_ = row + 1;
		cells_.resize(maxRows_, vector<BasicExcelCell>(maxCols_));
	}

	return &(cells_[row][col]);
}

// Erase content of a cell. row and col starts from 0.
// Returns true if successful, false if row or col exceeds range.
bool BasicExcelWorksheet::EraseCell(size_t row, size_t col)
{
	if (row<maxRows_ && col<maxCols_)
	{
		cells_[row][col].EraseContents();
		return true;
	}
	else return false;
}

// Update cells using information from BasicExcel.worksheets_.
void BasicExcelWorksheet::UpdateCells()
{
	// Define some reference
	Worksheet::Dimensions& dimension = excel_->worksheets_[sheetIndex_].dimensions_;
	vector<Worksheet::CellTable::RowBlock>& rRowBlocks = excel_->worksheets_[sheetIndex_].cellTable_.rowBlocks_;

	vector<wchar_t> wstr;
	vector<char> str;

	maxRows_ = dimension.lastUsedRowIndexPlusOne_;
	maxCols_ = dimension.lastUsedColIndexPlusOne_;

	// Resize the cells to the size of the worksheet
	vector<BasicExcelCell> cellCol(maxCols_);
	cells_.resize(maxRows_, cellCol);

	size_t maxRowBlocks = rRowBlocks.size();
	for (size_t i=0; i<maxRowBlocks; ++i)
	{
		vector<Worksheet::CellTable::RowBlock::CellBlock>& rCellBlocks = rRowBlocks[i].cellBlocks_;
		size_t maxCells = rCellBlocks.size();
		for (size_t j=0; j<maxCells; ++j)
		{
			size_t row = rCellBlocks[j].RowIndex();
			size_t col = rCellBlocks[j].ColIndex();
			switch (rCellBlocks[j].type_)
			{
				case CODE::BLANK:
					break;
			
				case CODE::BOOLERR:
					if (rCellBlocks[j].boolerr_.error_ == 0)
					{
						cells_[row][col].Set(rCellBlocks[j].boolerr_.code_);
					}
					break;
					
				case CODE::LABELSST:
				{
					vector<LargeString>& ss = excel_->workbook_.sst_.strings_;
					if (ss[rCellBlocks[j].labelsst_.SSTRecordIndex_].unicode_ & 1)
					{
						wstr = ss[rCellBlocks[j].labelsst_.SSTRecordIndex_].wname_;
						wstr.resize(wstr.size()+1);
						wstr.back() = L'\0';
						cells_[row][col].Set(&*(wstr.begin()));						
					}
					else
					{
						str = ss[rCellBlocks[j].labelsst_.SSTRecordIndex_].name_;
						str.resize(str.size()+1);
						str.back() = '\0';
						cells_[row][col].Set(&*(str.begin()));
					}
					break;
				}

				case CODE::MULBLANK:	
					break;
					
				case CODE::MULRK:
				{
					size_t maxCols = rCellBlocks[j].mulrk_.lastColIndex_ - rCellBlocks[j].mulrk_.firstColIndex_ + 1;
					for (size_t k=0; k<maxCols; ++k, ++col)
					{
						// Get values of the whole range
						int rkValue = rCellBlocks[j].mulrk_.XFRK_[k].RKValue_;
						if (IsRKValueAnInteger(rkValue))
						{
							cells_[row][col].Set(GetIntegerFromRKValue(rkValue));
						}
						else
						{
							cells_[row][col].Set(GetDoubleFromRKValue(rkValue));
						}
					}
					break;
				}

				case CODE::NUMBER:
					cells_[row][col].Set(rCellBlocks[j].number_.value_);
					break;

				case CODE::RK:
				{
					int rkValue = rCellBlocks[j].rk_.value_;
					if (IsRKValueAnInteger(rkValue))
					{
						cells_[row][col].Set(GetIntegerFromRKValue(rkValue));
					}
					else
					{
						cells_[row][col].Set(GetDoubleFromRKValue(rkValue));
					}
					break;
				}

				case CODE::FORMULA:
					break;
			}
		}
	}
}
/************************************************************************************************************/

/************************************************************************************************************/
BasicExcelCell::BasicExcelCell() : type_(UNDEFINED) {};

// Get type of value stored in current Excel cell. 
// Returns one of the enums.
int BasicExcelCell::Type() const {return type_;}

// Get an integer value.
// Returns false if cell does not contain an integer or a double.
bool BasicExcelCell::Get(int& val) const 
{
	if (type_ == INT)
	{
		val = ival_;
		return true;
	}
	else if (type_ == DOUBLE)
	{
		val = (int)dval_;
		return true;
	}
	else return false;
}

// Get a double value.
// Returns false if cell does not contain a double or an integer.
bool BasicExcelCell::Get(double& val) const 
{
	if (type_ == DOUBLE)
	{
		val = dval_;
		return true;
	}
	else if (type_ == INT)
	{
			val = (double)ival_;
			return true;
	}
	else return false;
}

// Get an ANSI string.
// Returns false if cell does not contain an ANSI string.
bool BasicExcelCell::Get(char* str) const 
{
	if (type_ == STRING)
	{
		if (str_.empty()) *str = '\0';
		else strcpy(str, &*(str_.begin()));
		return true;
	}
	else return false;
}

// Get an Unicode string.
// Returns false if cell does not contain an Unicode string.
bool BasicExcelCell::Get(wchar_t* str) const 
{
	if (type_ == WSTRING)
	{
		if (wstr_.empty()) *str = L'\0';
		else wcscpy(str, &*(wstr_.begin()));
		return true;
	}
	else return false;
}

// Return length of ANSI or Unicode string (excluding null character).
size_t BasicExcelCell::GetStringLength() const
{
	if (type_ == STRING) return str_.size() - 1;
	else return wstr_.size() - 1;
}

// Get an integer value.
// Returns 0 if cell does not contain an integer.
int BasicExcelCell::GetInteger() const
{
	int val;
	if (Get(val)) return val;
	else return 0;
}

// Get a double value.
// Returns 0.0 if cell does not contain a double.
double BasicExcelCell::GetDouble() const
{
	double val;
	if (Get(val)) return val;
	else return 0.0;
}

// Get an ANSI string.
// Returns 0 if cell does not contain an ANSI string.
const char* BasicExcelCell::GetString() const
{
	vector<char> str(str_.size());
	if (!str.empty() && Get(&*(str.begin()))) return &*(str_.begin());
	else return 0;
}

// Get an Unicode string.
// Returns 0 if cell does not contain an Unicode string.
const wchar_t* BasicExcelCell::GetWString() const
{
	vector<wchar_t> wstr(wstr_.size());
	if (!wstr.empty() && Get(&*(wstr.begin()))) return &*(wstr_.begin());
	else return 0;
}

// Set content of current Excel cell to an integer.
void BasicExcelCell::Set(int val) 
{
	SetInteger(val);
}

// Set content of current Excel cell to a double.
void BasicExcelCell::Set(double val) 
{
	SetDouble(val);
}

// Set content of current Excel cell to an ANSI string.
void BasicExcelCell::Set(const char* str) 
{
	SetString(str);
}

// Set content of current Excel cell to an Unicode string.
void BasicExcelCell::Set(const wchar_t* str)	
{
	SetWString(str);
}

// Set content of current Excel cell to an integer.
void BasicExcelCell::SetInteger(int val) 
{
	type_ = INT; 
	ival_ = val;
}

// Set content of current Excel cell to a double.
void BasicExcelCell::SetDouble(double val) 
{
	type_ = DOUBLE; 
	dval_ = val;
}

// Set content of current Excel cell to an ANSI string.
void BasicExcelCell::SetString(const char* str) 
{
	size_t length = strlen(str);
	if (length > 0)
	{
		type_ = STRING;
		str_ = vector<char>(length+1);
		strcpy(&*(str_.begin()), str);
		wstr_.clear();
	}
	else EraseContents();
}

// Set content of current Excel cell to an Unicode string.
void BasicExcelCell::SetWString(const wchar_t* str)	
{
	size_t length = wcslen(str);
	if (length > 0)
	{
		type_ = WSTRING;
		wstr_ = vector<wchar_t>(length+1);
		wcscpy(&*(wstr_.begin()), str);
		str_.clear();
	}
	else EraseContents();
}

// Erase the content of current Excel cell.
// Set type to UNDEFINED.
void BasicExcelCell::EraseContents()
{
	type_ = UNDEFINED;
	str_.clear();
	wstr_.clear();
}

///< Print cell to output stream.
///< Print a null character if cell is undefined.
ostream& operator<<(ostream& os, const BasicExcelCell& cell)
{
	switch (cell.Type())
	{
		case BasicExcelCell::UNDEFINED:
			os << '\0';
			break;

		case BasicExcelCell::INT:
			os << cell.GetInteger();
			break;

		case BasicExcelCell::DOUBLE:
			os << cell.GetDouble();
			break;

		case BasicExcelCell::STRING:
			os << cell.GetString();
			break;

		case BasicExcelCell::WSTRING:
			os << cell.GetWString();
			break;
	}
	return os;
}

}
using namespace YExcel;
typedef unsigned char byte;
typedef unsigned int uint32;

using std::string;
using std::ifstream;

//MD5();
//MD5(const void* input, size_t length);
//MD5(const string& str);
//MD5(ifstream& in);
void update(const void* input, size_t length);
void update(const string& str);
void update(ifstream& in);
const byte* digest();
void reset();

void update(const byte* input, size_t length);
void final();
void transform(const byte block[64]);
void encode(const uint32* input, byte* output, size_t length);
void decode(const byte* input, uint32* output, size_t length);
string bytesToHexString(const byte* input, size_t length);

uint32 _state[4]; /* state (ABCD) */
uint32 _count[2]; /* number of bits, modulo 2^64 (low-order word first) */
byte _buffer[64]; /* input buffer */
byte _digest[16]; /* message digest */
bool _finished;   /* calculate finished ? */

enum { BUFFER_SIZE = 1024 };

/* Constants for MD5Transform routine. */
#define S11 7
#define S12 12
#define S13 17
#define S14 22
#define S21 5
#define S22 9
#define S23 14
#define S24 20
#define S31 4
#define S32 11
#define S33 16
#define S34 23
#define S41 6
#define S42 10
#define S43 15
#define S44 21


/* F, G, H and I are basic MD5 functions.
*/
#define F(x, y, z) (((x) & (y)) | ((~x) & (z)))
#define G(x, y, z) (((x) & (z)) | ((y) & (~z)))
#define H(x, y, z) ((x) ^ (y) ^ (z))
#define I(x, y, z) ((y) ^ ((x) | (~z)))

/* ROTATE_LEFT rotates x left n bits.
*/
#define ROTATE_LEFT(x, n) (((x) << (n)) | ((x) >> (32-(n))))

/* FF, GG, HH, and II transformations for rounds 1, 2, 3, and 4.
Rotation is separate from addition to prevent recomputation.
*/
#define FF(a, b, c, d, x, s, ac) { \
	(a) += F ((b), (c), (d)) + (x) + ac; \
	(a) = ROTATE_LEFT ((a), (s)); \
	(a) += (b); \
}
#define GG(a, b, c, d, x, s, ac) { \
	(a) += G ((b), (c), (d)) + (x) + ac; \
	(a) = ROTATE_LEFT ((a), (s)); \
	(a) += (b); \
}
#define HH(a, b, c, d, x, s, ac) { \
	(a) += H ((b), (c), (d)) + (x) + ac; \
	(a) = ROTATE_LEFT ((a), (s)); \
	(a) += (b); \
}
#define II(a, b, c, d, x, s, ac) { \
	(a) += I ((b), (c), (d)) + (x) + ac; \
	(a) = ROTATE_LEFT ((a), (s)); \
	(a) += (b); \
}


static const byte PADDING[64] = { 0x80 };
static const char HEX[16] = {
	'0', '1', '2', '3',
		'4', '5', '6', '7',
		'8', '9', 'a', 'b',
		'c', 'd', 'e', 'f'
};


/* Return the message-digest */
const byte* digest() {
	
	if (!_finished) {
		_finished = true;
		final();
	}
	return _digest;
}

/* Reset the calculate state */
void reset() {
	
	_finished = false;
	/* reset number of bits. */
	_count[0] = _count[1] = 0;
	/* Load magic initialization constants. */
	_state[0] = 0x67452301;
	_state[1] = 0xefcdab89;
	_state[2] = 0x98badcfe;
	_state[3] = 0x10325476;
}

/* Updating the context with a input buffer. */
void update(const void* input, size_t length) {
	update((const byte*)input, length);
}

/* Updating the context with a string. */
void update(const string& str) {
	update((const byte*)str.c_str(), str.length());
}

/* Updating the context with a file. */
void update(ifstream& in) {
	if(in)
	{
		std::streamsize length;
		char buffer[BUFFER_SIZE];
		while (!in.eof()) {
			in.read(buffer, BUFFER_SIZE);
			length = in.gcount();
			if (length > 0) {
				update(buffer, length);
			}
		}
		in.close();
	}
}

/* MD5 block update operation. Continues an MD5 message-digest
operation, processing another message block, and updating the
context.
*/
void update(const byte* input, size_t length) {
	
	uint32 i, index, partLen;
	
	_finished = false;
	
	/* Compute number of bytes mod 64 */
	index = (uint32)((_count[0] >> 3) & 0x3f);
	
	/* update number of bits */
	if ((_count[0] += ((uint32)length << 3)) < ((uint32)length << 3)) {
		++_count[1];
	}
	_count[1] += ((uint32)length >> 29);
	
	partLen = 64 - index;
	
	/* transform as many times as possible. */
	if (length >= partLen) {
		
		memcpy(&_buffer[index], input, partLen);
		transform(_buffer);
		
		for (i = partLen; i + 63 < length; i += 64) {
			transform(&input[i]);
		}
		index = 0;
		
	} else {
		i = 0;
	}
	
	/* Buffer remaining input */
	memcpy(&_buffer[index], &input[i], length - i);
}

/* MD5 finalization. Ends an MD5 message-_digest operation, writing the
the message _digest and zeroizing the context.
*/
void final() {
	
	byte bits[8];
	uint32 oldState[4];
	uint32 oldCount[2];
	uint32 index, padLen;
	
	/* Save current state and count. */
	memcpy(oldState, _state, 16);
	memcpy(oldCount, _count, 8);
	
	/* Save number of bits */
	encode(_count, bits, 8);
	
	/* Pad out to 56 mod 64. */
	index = (uint32)((_count[0] >> 3) & 0x3f);
	padLen = (index < 56) ? (56 - index) : (120 - index);
	update(PADDING, padLen);
	
	/* Append length (before padding) */
	update(bits, 8);
	
	/* Store state in digest */
	encode(_state, _digest, 16);
	
	/* Restore current state and count. */
	memcpy(_state, oldState, 16);
	memcpy(_count, oldCount, 8);
}

/* MD5 basic transformation. Transforms _state based on block. */
void transform(const byte block[64]) {
	
	uint32 a = _state[0], b = _state[1], c = _state[2], d = _state[3], x[16];
	
	decode(block, x, 64);
	
	/* Round 1 */
	FF (a, b, c, d, x[ 0], S11, 0xd76aa478); /* 1 */
	FF (d, a, b, c, x[ 1], S12, 0xe8c7b756); /* 2 */
	FF (c, d, a, b, x[ 2], S13, 0x242070db); /* 3 */
	FF (b, c, d, a, x[ 3], S14, 0xc1bdceee); /* 4 */
	FF (a, b, c, d, x[ 4], S11, 0xf57c0faf); /* 5 */
	FF (d, a, b, c, x[ 5], S12, 0x4787c62a); /* 6 */
	FF (c, d, a, b, x[ 6], S13, 0xa8304613); /* 7 */
	FF (b, c, d, a, x[ 7], S14, 0xfd469501); /* 8 */
	FF (a, b, c, d, x[ 8], S11, 0x698098d8); /* 9 */
	FF (d, a, b, c, x[ 9], S12, 0x8b44f7af); /* 10 */
	FF (c, d, a, b, x[10], S13, 0xffff5bb1); /* 11 */
	FF (b, c, d, a, x[11], S14, 0x895cd7be); /* 12 */
	FF (a, b, c, d, x[12], S11, 0x6b901122); /* 13 */
	FF (d, a, b, c, x[13], S12, 0xfd987193); /* 14 */
	FF (c, d, a, b, x[14], S13, 0xa679438e); /* 15 */
	FF (b, c, d, a, x[15], S14, 0x49b40821); /* 16 */
	
	/* Round 2 */
	GG (a, b, c, d, x[ 1], S21, 0xf61e2562); /* 17 */
	GG (d, a, b, c, x[ 6], S22, 0xc040b340); /* 18 */
	GG (c, d, a, b, x[11], S23, 0x265e5a51); /* 19 */
	GG (b, c, d, a, x[ 0], S24, 0xe9b6c7aa); /* 20 */
	GG (a, b, c, d, x[ 5], S21, 0xd62f105d); /* 21 */
	GG (d, a, b, c, x[10], S22, 0x2441453); /* 22 */
	GG (c, d, a, b, x[15], S23, 0xd8a1e681); /* 23 */
	GG (b, c, d, a, x[ 4], S24, 0xe7d3fbc8); /* 24 */
	GG (a, b, c, d, x[ 9], S21, 0x21e1cde6); /* 25 */
	GG (d, a, b, c, x[14], S22, 0xc33707d6); /* 26 */
	GG (c, d, a, b, x[ 3], S23, 0xf4d50d87); /* 27 */
	GG (b, c, d, a, x[ 8], S24, 0x455a14ed); /* 28 */
	GG (a, b, c, d, x[13], S21, 0xa9e3e905); /* 29 */
	GG (d, a, b, c, x[ 2], S22, 0xfcefa3f8); /* 30 */
	GG (c, d, a, b, x[ 7], S23, 0x676f02d9); /* 31 */
	GG (b, c, d, a, x[12], S24, 0x8d2a4c8a); /* 32 */
	
	/* Round 3 */
	HH (a, b, c, d, x[ 5], S31, 0xfffa3942); /* 33 */
	HH (d, a, b, c, x[ 8], S32, 0x8771f681); /* 34 */
	HH (c, d, a, b, x[11], S33, 0x6d9d6122); /* 35 */
	HH (b, c, d, a, x[14], S34, 0xfde5380c); /* 36 */
	HH (a, b, c, d, x[ 1], S31, 0xa4beea44); /* 37 */
	HH (d, a, b, c, x[ 4], S32, 0x4bdecfa9); /* 38 */
	HH (c, d, a, b, x[ 7], S33, 0xf6bb4b60); /* 39 */
	HH (b, c, d, a, x[10], S34, 0xbebfbc70); /* 40 */
	HH (a, b, c, d, x[13], S31, 0x289b7ec6); /* 41 */
	HH (d, a, b, c, x[ 0], S32, 0xeaa127fa); /* 42 */
	HH (c, d, a, b, x[ 3], S33, 0xd4ef3085); /* 43 */
	HH (b, c, d, a, x[ 6], S34, 0x4881d05); /* 44 */
	HH (a, b, c, d, x[ 9], S31, 0xd9d4d039); /* 45 */
	HH (d, a, b, c, x[12], S32, 0xe6db99e5); /* 46 */
	HH (c, d, a, b, x[15], S33, 0x1fa27cf8); /* 47 */
	HH (b, c, d, a, x[ 2], S34, 0xc4ac5665); /* 48 */
	
	/* Round 4 */
	II (a, b, c, d, x[ 0], S41, 0xf4292244); /* 49 */
	II (d, a, b, c, x[ 7], S42, 0x432aff97); /* 50 */
	II (c, d, a, b, x[14], S43, 0xab9423a7); /* 51 */
	II (b, c, d, a, x[ 5], S44, 0xfc93a039); /* 52 */
	II (a, b, c, d, x[12], S41, 0x655b59c3); /* 53 */
	II (d, a, b, c, x[ 3], S42, 0x8f0ccc92); /* 54 */
	II (c, d, a, b, x[10], S43, 0xffeff47d); /* 55 */
	II (b, c, d, a, x[ 1], S44, 0x85845dd1); /* 56 */
	II (a, b, c, d, x[ 8], S41, 0x6fa87e4f); /* 57 */
	II (d, a, b, c, x[15], S42, 0xfe2ce6e0); /* 58 */
	II (c, d, a, b, x[ 6], S43, 0xa3014314); /* 59 */
	II (b, c, d, a, x[13], S44, 0x4e0811a1); /* 60 */
	II (a, b, c, d, x[ 4], S41, 0xf7537e82); /* 61 */
	II (d, a, b, c, x[11], S42, 0xbd3af235); /* 62 */
	II (c, d, a, b, x[ 2], S43, 0x2ad7d2bb); /* 63 */
	II (b, c, d, a, x[ 9], S44, 0xeb86d391); /* 64 */
	
	_state[0] += a;
	_state[1] += b;
	_state[2] += c;
	_state[3] += d;
}

/* Encodes input (ulong) into output (byte). Assumes length is
a multiple of 4.
*/
void encode(const uint32* input, byte* output, size_t length) {
	
	for (size_t i = 0, j = 0; j < length; ++i, j += 4) {
		output[j]= (byte)(input[i] & 0xff);
		output[j + 1] = (byte)((input[i] >> 8) & 0xff);
		output[j + 2] = (byte)((input[i] >> 16) & 0xff);
		output[j + 3] = (byte)((input[i] >> 24) & 0xff);
	}
}

/* Decodes input (byte) into output (ulong). Assumes length is
a multiple of 4.
*/
void decode(const byte* input, uint32* output, size_t length) {
	
	for (size_t i = 0, j = 0; j < length; ++i, j += 4) {
		output[i] = ((uint32)input[j]) | (((uint32)input[j + 1]) << 8) |
			(((uint32)input[j + 2]) << 16) | (((uint32)input[j + 3]) << 24);
	}
}

/* Convert byte array to hex string. */
string bytesToHexString(const byte* input, size_t length) {
	
	string str;
	str.reserve(length << 1);
	for (size_t i = 0; i < length; ++i) {
		int t = input[i];
		int a = t / 16;
		int b = t % 16;
		str.append(1, HEX[a]);
		str.append(1, HEX[b]);
	}
	return str;
}
bool Exists(string fileName) {
    struct _stat buf;
    return _stat(fileName.c_str(), &buf) == 0;
}
int main(int argc, char* argv[])
{
	BasicExcel e;
	// create a workbook
	e.New();
	// create a sheet
	BasicExcelWorksheet* sheet = e.AddWorksheet("PascalTriangle");
	const int len=10;
	int a[len][len];
	int g,i,j,k,m,n;


//分配第一列为1
	for (i=0;i<len;i++)
	{
		a[i][0]=1;
	}
//分配斜边为1
	for (g=0;g<len;g++)
	{
		a[g][g]=1;
	}

///////////////////////////////////////////////////////////

//杨辉三角赋值
	for (m=1;m<len;m++)//列扫深度len-1
	{
		for (n=1;n<m;n++)//行扫宽度m-1
		{
			a[m][n]=a[m-1][n-1]+a[m-1][n];
		}
	}

//输出
	for (j=0;j<len;j++)//扫列
	{
		for (k=0;k<=j;k++)//扫行宽度j
			sheet->Cell(j,k)->SetInteger(a[j][k]);
	}
	TCHAR xlsname[] = "PascalTriangle";
	TCHAR xlsfile[FILENAME_MAX];
	sprintf(xlsfile, "%s.xls", xlsname);	
	if(Exists(xlsfile))
	{
		int num = 2;
		while(Exists(xlsfile))
			sprintf(xlsfile, "%s%d.xls", xlsname, num++);		
	}
	e.SaveAs(xlsfile);
	return 0;
}

//int main(int argc, char** argv) {
//	return 0;
//}
