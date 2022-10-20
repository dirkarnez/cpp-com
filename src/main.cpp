// C:\Program Files\PDFCreator\Com Scripts\C++\CppCOMTest\CppCOMTest\CppCOMTest.cpp

/*
 * PDFCreator COM Interface tests for C++
 * Part of the PDFCreator application
 * License: GPL
 * Homepage: http://www.pdfforge.org/pdfcreator
 * Version: 1.1.0.0
 * Created: May, 25. 2020
 * Modified: May, 25. 2020
 * Author: pdfforge GmbH
 * Comments: This project demonstrates the use of the COM Interface of PDFCreator.
			 There are 3 different kinds of usage presented.
			 Further usage presentation is only available in the JavaScript directory.
 * Note: When executed in release mode paths have to be modified.
 */

#include "stdafx.h"
#include <iostream>
#include <sstream>
#include <ShlObj.h>
#include <stdlib.h>
#include <string>
#include <algorithm>
#include "comdef.h"
#import "C:\Program Files\PDFCreator\PDFCreator.COM.tlb"		//IMPORTANT: PUT THE PATH TO THE PDFCreator.tlb HERE, IT MIGHT DIFFER FROM THE PROVIDED ONE

using namespace PDFCreator_COM;

//Function declarations
void testPage2Pdf(IQueuePtr qPtr);
void testPage2Jpeg(IQueuePtr qPtr);
void testPagesMerged2Tif(IQueuePtr qPtr);

int _tmain(int argc, _TCHAR* argv[])
{
	::CoInitialize(NULL);
	IQueuePtr qPtr(__uuidof(Queue));

	try
	{
		if (qPtr != NULL)
		{
			unsigned int choice = 0;

			//Initializing the queue object
			qPtr->Initialize();

			std::cout << "What do you want to do?" << std::endl << std::endl;
			std::cout << "1. Convert windows testpage to a pdf file." << std::endl;
			std::cout << "2. Convert windows testpage to a jpeg file with settings changes." << std::endl;
			std::cout << "3. Print several windows testpages, merge them and convert them to a tif file." << std::endl;

			do
			{
				std::cout << "Please enter a number between 1 and 3 corresponding to the above choices or 9 for exiting: ";
				std::cin >> choice;

				switch (choice) {
				case 1:
					testPage2Pdf(qPtr);
					break;
				case 2:
					testPage2Jpeg(qPtr);
					break;
				case 3:
					testPagesMerged2Tif(qPtr);
					break;
				case 9:
					break;
				default:
					std::cout << "Illegal number." << std::endl;
				}
			} while (choice != 9);

			qPtr->ReleaseCom();
		}
		else
		{
			std::cout << "Could not create a pointer to the queue object. Please enter a key to exit." << std::endl;
			getchar();
			::CoUninitialize();
			return -1;
		}
	}
	catch (...)
	{
		std::cout << "Something went wrong. Releasing the COM object." << std::endl;
		getchar();
		qPtr->ReleaseCom();
		::CoUninitialize();
		return -1;
	}

	::CoUninitialize();
	return 0;
}

void printTestPage()
{
	SHELLEXECUTEINFO ShExecInfo;
	ShExecInfo.cbSize = sizeof(SHELLEXECUTEINFO);
	ShExecInfo.fMask = NULL;
	ShExecInfo.hwnd = NULL;
	ShExecInfo.lpVerb = TEXT("open");
	ShExecInfo.lpFile = TEXT("RUNDLL32.exe");
	ShExecInfo.lpParameters = TEXT("PRINTUI.DLL,PrintUIEntry /k /n \"PDFCreator\"");
	ShExecInfo.lpDirectory = NULL;
	ShExecInfo.nShow = SW_NORMAL;
	ShExecInfo.hInstApp = NULL;

	std::cout << "Printing a windows test page..." << std::endl;
	ShellExecuteEx(&ShExecInfo);
}

bool replace(std::string& str, const std::string& from, const std::string& to) {
	size_t start_pos = str.find(from);
	if (start_pos == std::string::npos)
		return false;
	str.replace(start_pos, from.length(), to);
	return true;
}

std::string createFilePath(std::string dir, std::string fileName)
{
	wchar_t wPath[MAX_PATH];
	std::stringstream ss;

	ss << dir << fileName;

	// Will contain exe path
	HMODULE hModule = GetModuleHandle(NULL);
	if (hModule != NULL)
	{
		// When passing NULL to GetModuleHandle, it returns handle of exe itself
		GetModuleFileName(hModule, wPath, MAX_PATH);
	}
	std::wstring ws(wPath);
	std::string str(ws.begin(), ws.end());

	replace(str, "Debug\\CppCOMTest.exe", ss.str());

	std::cout << str << std::endl;

	return str;
}

BSTR createFileName(std::string fileNameExt)
{
	std::stringstream ss;
	ss << "Testpage_ " << fileNameExt << ".pdf";
	std::string fileName(createFilePath("Results\\", ss.str()));
	std::wstring ws(fileName.begin(), fileName.end());
	BSTR bs = SysAllocStringLen(ws.data(), ws.size());

	return bs;
}

BSTR StrToBstr(std::string str)
{
	std::wstring ws(str.begin(), str.end());
	BSTR bs = SysAllocStringLen(ws.data(), ws.size());

	return bs;
}

void testPage2Pdf(IQueuePtr qPtr)
{
	printTestPage();

	if (!qPtr->WaitForJob(10))
	{
		std::cout << "The print job did not reach the queue within 10 seconds" << std::endl;
	}
	else
	{
		std::stringstream ss;
		ss << "Currently there are " << qPtr->Count << " job(s) in the queue.";
		std::cout << ss.str() << std::endl;

		std::cout << "Getting job instance" << std::endl;
		IPrintJobPtr printJob = qPtr->NextJob;

		printJob->SetProfileByGuid("DefaultGuid");

		std::cout << "Converting under DefaultGuid conversion profile" << std::endl;
		printJob->ConvertTo(createFileName("2PDF"));

		if (!printJob->IsFinished || !printJob->IsSuccessful)
		{
			ss.flush();
			ss << "Could not convert the file: " << createFileName("2PDF");
			std::cout << ss.str() << std::endl;
		}
		else
		{
			std::cout << "Job finished successfully" << std::endl;
		}
	}
}

void testPage2Jpeg(IQueuePtr qPtr)
{
	printTestPage();

	if (!qPtr->WaitForJob(10))
	{
		std::cout << "The print job did not reach the queue within 10 seconds" << std::endl;
	}
	else
	{
		std::stringstream ss;
		ss << "Currently there are " << qPtr->Count << " job(s) in the queue.";
		std::cout << ss.str() << std::endl;

		std::cout << "Getting job instance" << std::endl;
		IPrintJobPtr printJob = qPtr->NextJob;

		printJob->SetProfileByGuid("JpegGuid");

		//The SetProfileSettings method allows us to change the JpegSettings of the job
		//We want 24 bit colors for our converted file
		printJob->SetProfileSetting("JpegSettings.Color", "Color24Bit");
		printJob->SetProfileSetting("JpegSettings.Quality", "100");

		std::cout << "Converting under JpegtGuid conversion profile" << std::endl;
		printJob->ConvertTo(createFileName("2Jpg"));

		if (!printJob->IsFinished || !printJob->IsSuccessful)
		{
			ss.flush();
			ss << "Could not convert the file: " << createFileName("2Jpg");
			std::cout << ss.str() << std::endl;
		}
		else
		{
			std::cout << "Job finished successfully" << std::endl;
		}
	}
}

void testPagesMerged2Tif(IQueuePtr qPtr)
{
	printTestPage();
	printTestPage();
	printTestPage();

	if (!qPtr->WaitForJobs(3, 15))
	{
		std::cout << "The print jobs did not reach the queue within 15 seconds" << std::endl;
	}
	else
	{
		std::stringstream ss;
		ss << "Currently there are " << qPtr->Count << " job(s) in the queue.";
		std::cout << ss.str() << std::endl;
		std::cout << "Merging all available jobs." << std::endl;
		qPtr->MergeAllJobs();

		std::cout << "Getting job instance" << std::endl;
		IPrintJobPtr printJob = qPtr->NextJob;

		printJob->SetProfileByGuid("DefaultGuid");

		std::cout << "Setting output format of DefaultGuid to tif." << std::endl;
		printJob->SetProfileSetting("OutputFormat", "Tif");

		std::cout << "Converting under DefaultGuid conversion profile but with tif as output format" << std::endl;
		printJob->ConvertTo(createFileName("Merged2Tif"));

		if (!printJob->IsFinished || !printJob->IsSuccessful)
		{
			ss.flush();
			ss << "Could not convert the file: " << createFileName("Merged2Tif");
			std::cout << ss.str() << std::endl;
		}
		else
		{
			std::cout << "Job finished successfully" << std::endl;
		}
	}
}
