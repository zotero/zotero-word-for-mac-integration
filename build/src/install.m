/*
	 ***** BEGIN LICENSE BLOCK *****
	 
	 Copyright (c) 2011  Zotero
	 Center for History and New Media
	 George Mason University, Fairfax, Virginia, USA
	 http://zotero.org
	 
	 Zotero is free software: you can redistribute it and/or modify
	 it under the terms of the GNU Affero General Public License as published by
	 the Free Software Foundation, either version 3 of the License, or
	 (at your option) any later version.
	 
	 Zotero is distributed in the hope that it will be useful,
	 but WITHOUT ANY WARRANTY; without even the implied warranty of
	 MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
	 GNU Affero General Public License for more details.
	 
	 You should have received a copy of the GNU Affero General Public License
	 along with Zotero.  If not, see <http://www.gnu.org/licenses/>.
	 
	 ***** END LICENSE BLOCK *****
 */

#include <copyfile.h>
#include "zoteroMacWordIntegration.h"
#include "Finder.h"

#define AUTH_PROMPT "Zotero Word for Mac Integration requires administator "\
	"privileges to complete installation."
#define PATH_MKDIR "/bin/mkdir"
#define PATH_RM "/bin/rm"
#define PATH_DITTO "/usr/bin/ditto"

statusCode installTemplate(NSString* templatePath, NSString* path);
statusCode installScripts(NSString* templatePath);
statusCode parseAuthorizationStatus(OSStatus status, const char file[],
									const char function[], unsigned int line);
statusCode getAuthorizationRef(NSString* templatePath);
statusCode performAuthorizedMkdir(NSString* templatePath, NSString *path,
								  BOOL emptyIfExists);
statusCode performAuthorizedCopy(NSString* templatePath, NSString* path1,
								 NSString* path2);

statusCode install(const char templatePath[]) {
	NSString* templatePathNS = [NSString stringWithUTF8String:templatePath];
	FinderApplication* finder = nil;
	
	// Fix template type and creator
	NSError *err = nil;
	NSFileManager *fm = [NSFileManager defaultManager];
	if(![fm setAttributes:[NSDictionary dictionaryWithObjectsAndKeys:
					   [NSNumber numberWithUnsignedInt:'MSWD'],
					   NSFileHFSCreatorCode,
					   [NSNumber numberWithUnsignedInt:'W8TN'],
					   NSFileHFSTypeCode, nil]
			 ofItemAtPath:templatePathNS error:&err]) {
		DIE([err localizedDescription]);
	}
	
	// Get directories
	NSArray *temp = NSSearchPathForDirectoriesInDomains(NSApplicationDirectory,
														NSLocalDomainMask, YES);
	if(!temp || ![temp count]) DIE(@"Could not find applications directory");
	NSString* applicationsFolder = [temp objectAtIndex:0];
	
	// Look for Word in obvious places
	NSMutableSet* wordLocations = [NSMutableSet setWithCapacity:3];
	[wordLocations addObject:[applicationsFolder stringByAppendingPathComponent:
							  @"Microsoft Office 2004/Microsoft Word"]];
	[wordLocations addObject:[applicationsFolder stringByAppendingPathComponent:
							  @"Microsoft Office 2008/Microsoft Word.app"]];
	[wordLocations addObject:[applicationsFolder stringByAppendingPathComponent:
							  @"Microsoft Office 2011/Microsoft Word.app"]];
	
	// Look for Word bundle
	NSWorkspace *workspace = [NSWorkspace sharedWorkspace];
	NSString* testPath = [workspace absolutePathForAppBundleWithIdentifier:
						  @"com.Microsoft.Word"];
	if(testPath) [wordLocations addObject:testPath];
	
	// Look for Word by creator code
	CFURLRef testURL;
	if(LSFindApplicationForInfo('MSWD', NULL, NULL, NULL, &testURL) && testURL) {
		[wordLocations addObject:[(NSURL*)testURL path]];
	}
	
	BOOL shouldInstallScripts = NO, installed = NO, wordXFound = NO;
	for(NSString* path in wordLocations) {
		BOOL isDirectory;
		NSString* version;
		
		// Obviously, skip if app doesn't exist
		if(![fm fileExistsAtPath:path isDirectory:&isDirectory]) continue;
		if(isDirectory) {
			// A bundle, so get the version using NSBundle
			version = [[[NSBundle bundleWithPath:path] infoDictionary]
								 objectForKey:(NSString*)kCFBundleVersionKey];
		} else {
			// Use Finder to get version of CFM app
			if(!finder) {
				finder = [SBApplication applicationWithBundleIdentifier:
						  @"com.apple.Finder"];
			}
			FinderFile* file = [[finder files]
								objectWithName:posixPathToHFSPath(path)];
			version = [file version];
		}
		
		if(!version) continue;
		NSUInteger dotIndex = [version rangeOfString:@"."].location;
		if(dotIndex == NSNotFound) continue;
		
		NSInteger intVersion = [[version substringToIndex:dotIndex] intValue];
		if(intVersion == 10) {
			wordXFound = YES;
		}
		
		// Install scripts for Word 2008 or later
		shouldInstallScripts = installScripts || intVersion >= 12;
		// Install scripts for Word 2004 or Word 2011
		if(intVersion == 11 || intVersion >= 14) {
			statusCode status = installTemplate(templatePathNS, path);
			if(status) return status;
			installed = true;
		}
	}
	
	if(installScripts) {
		statusCode status = installScripts(templatePathNS);
		if(status) return status;
		installed = true;
	}
	
	if(!installed) {
		if(wordXFound) {
			DIE(@"While an installation of Word X was found, Zotero Word for "
				"Mac Integration requires Word 2004 or later. Please upgrade "
				"Word to use Zotero Word Integration.");
		} else {
			DIE(@"Neither Word 2004 nor Word 2008 appear to be installed on "
				"this computer. Please install Microsoft Word, then try again.");
		}
	}
	
	// Check if Word is running
	BOOL wordIsRunning = NO;
	if([workspace respondsToSelector:@selector(runningApplications)]) {
		NSArray* apps = [workspace runningApplications];
		for(NSRunningApplication* app in apps) {
			NSRange range = [[app localizedName] rangeOfString:@"Microsoft Word"];
			if(range.location != NSNotFound) {
				wordIsRunning = YES;
				break;
			}
		}
	} else {
		NSArray* apps = [workspace launchedApplications];
		for(NSDictionary* app in apps) {
			NSRange range = [[app objectForKey:@"NSApplicationName"] 
							 rangeOfString:@"Microsoft Word"];
			if(range.location != NSNotFound) {
				wordIsRunning = YES;
				break;
			}
		}
	}
	
	if(wordIsRunning) {
		NSAlert *alert = [NSAlert alertWithMessageText:@"Zotero Word for Mac "
						  "Integration has been successfully installed, but "
						  "Word must be restarted before it can be used."
										 defaultButton:nil
									   alternateButton:nil
										   otherButton:nil
							 informativeTextWithFormat:@"Please restart Word "
						  "before continuing."];
		[alert runModal];
	}
	
	return STATUS_OK;
}

statusCode installTemplate(NSString* templatePath, NSString* path) {
	// Get path to startup folder
	NSString* officeDirectory = [path stringByDeletingLastPathComponent];
	NSString* startupDirectory = [officeDirectory stringByAppendingPathComponent:
								  @"Office/Startup/Word"];
	
	BOOL isDirectory = NO;
	NSFileManager *fm = [NSFileManager defaultManager];
	if(![fm fileExistsAtPath:startupDirectory isDirectory:&isDirectory]) {
		// If no Startup directory, look for a Start directory
		NSString* testStartupDirectory = [officeDirectory
										  stringByAppendingPathComponent:
										  @"Office/Start/Word"];
		if([fm fileExistsAtPath:testStartupDirectory
					 isDirectory:&isDirectory]) {
			if(!isDirectory) {
				NSString* err = [NSString
								 stringWithFormat:@"%@ is not a directory",
								 testStartupDirectory];
				DIE(err);
			}
			startupDirectory = testStartupDirectory;
		}
	}
	
	// Make sure directory exists, but don't empty it
	statusCode status = performAuthorizedMkdir(templatePath, startupDirectory,
											   false);
	if(status) return status;
	
	// Try to copy template to directory
	status = performAuthorizedCopy(templatePath, templatePath,
								   [startupDirectory
									stringByAppendingPathComponent:
									@"Zotero.dot"]
								   );
	if(status) return status;
	
	return STATUS_OK;
}

statusCode installScripts(NSString* templatePath) {
	NSString* potentialMicrosoftUserDataNames[] = {@"Microsoft User Data",
		@"Microsoft-Benutzerdaten", @"Données Utilisateurs Microsoft",
		@"Datos del Usuario de Microsoft", @"Datos de Usuario de Microsoft",
		@"Dati utente Microsoft", NULL};
	NSString* microsoftUserDataFolder = nil;
	NSFileManager *fm = [NSFileManager defaultManager];
	
	NSArray *temp = NSSearchPathForDirectoriesInDomains(NSDocumentDirectory,
														NSUserDomainMask, YES);
	if(!temp || ![temp count]) DIE(@"Could not find documents directory");
	NSString* documentsFolder = [temp objectAtIndex:0];
	
	temp = NSSearchPathForDirectoriesInDomains(NSLibraryDirectory,
											   NSUserDomainMask, YES);
	if(!temp || ![temp count]) DIE(@"Could not find library directory");
	NSString* preferencesFolder = [[temp objectAtIndex:0]
								   stringByAppendingPathComponent:
								   @"Preferences"];
	
	// Look for Script Menu Items folder in various places
	for(unsigned short i=0;
		!microsoftUserDataFolder && potentialMicrosoftUserDataNames[i] != NULL;
		i++) {
		NSString *testPath = [[documentsFolder stringByAppendingPathComponent:
							   potentialMicrosoftUserDataNames[i]]
							  stringByAppendingPathComponent:
							  @"Word Script Menu Items"];
		BOOL isDirectory;
		if([fm fileExistsAtPath:testPath isDirectory:&isDirectory]
		   && isDirectory) {
			microsoftUserDataFolder = testPath;
			break;
		} else {
			testPath = [[preferencesFolder stringByAppendingPathComponent:
						 potentialMicrosoftUserDataNames[i]]
						stringByAppendingPathComponent:
						@"Word Script Menu Items"];
			if([fm fileExistsAtPath:testPath isDirectory:&isDirectory]
			   && isDirectory) {
				microsoftUserDataFolder = testPath;
				break;
			}
		}
	}
	
	// We couldn't find Script Menu Items folder, so ask user to locate it
	if(!microsoftUserDataFolder) {
		NSOpenPanel* openPanel = [NSOpenPanel openPanel];
		[openPanel setCanChooseFiles:NO];
		[openPanel setCanChooseDirectories:YES];
		[openPanel setAllowsMultipleSelection:NO];
		[openPanel setMessage:@"Select the Word Script Menu Items folder, "
		 "usually located in Documents/Microsoft User Data"];
		[openPanel setDirectory:documentsFolder];
		if([openPanel runModal] == NSFileHandlingPanelOKButton) {
			microsoftUserDataFolder = [[[openPanel URLs] objectAtIndex:0] path];
		} else {
			return STATUS_EXCEPTION_ALREADY_DISPLAYED;
		}
	}
	
	// Clobber the Zotero folder
	NSString* zoteroFolder = [microsoftUserDataFolder
							  stringByAppendingPathComponent:@"Zotero"];
	statusCode status = performAuthorizedMkdir(templatePath, zoteroFolder,
											   true);
	if(status) return status;
	
	// Generate the scripts
	const char* W2008_SCRIPT_NAMES[] = {"Add Bibliography\\cob.scpt",
		"Add Citation\\coa.scpt", "Edit Bibliography\\cod.scpt",
		"Edit Citation\\coe.scpt", "Refresh\\cor.scpt",
		"Remove Field Codes.scpt", "Set Document Preferences\\cop.scpt",
		NULL};
	const char* W2008_SCRIPT_COMMANDS[] = {"addBibliography", "addCitation",
		"editBibliography", "editCitation", "refresh", "removeCodes",
		"setDocPrefs", NULL};
	for(unsigned short i=0; W2008_SCRIPT_NAMES[i] != NULL; i++) {
		NSTask *task;
		task = [[NSTask alloc] init];
		[task autorelease];
		[task setLaunchPath: @"/usr/bin/osacompile"];
		NSPipe *outPipe = [NSPipe pipe];
		NSPipe *inPipe = [NSPipe pipe];
		[task setStandardInput:inPipe];
		[task setStandardOutput:outPipe];
		NSFileHandle *outFile = [outPipe fileHandleForReading];
		NSFileHandle *inFile = [inPipe fileHandleForWriting];
		
		SInt32 versMaj, versMin;
		Gestalt(gestaltSystemVersionMajor, &versMaj);
		Gestalt(gestaltSystemVersionMinor, &versMin);
		NSString* scriptPath = [zoteroFolder stringByAppendingPathComponent:
								[NSString stringWithUTF8String: 
								 W2008_SCRIPT_NAMES[i]]];
		if(versMaj > 10 || versMin >= 7) {
			// Lion doesn't add a type or creator by default
			[task setArguments:[NSArray arrayWithObjects: @"-t", @"osas", @"-c",
								@"ToyS", @"-o", scriptPath, nil]];
		} else {
			// Older versions of Mac OS X support the arguments above, but they
			// somehow get reversed.
			[task setArguments:[NSArray arrayWithObjects: @"-o", scriptPath,
								nil]];
		}
		
		[task launch];
		[inFile writeData:[[NSString stringWithFormat:@"try\n"
		 "do shell script \"PIPE=\\\"/Users/Shared/.zoteroIntegrationPipe_$LOGNAME\\\";  if [ ! -e \\\"$PIPE\\\" ]; then PIPE=~/.zoteroIntegrationPipe; fi; if [ -e \\\"$PIPE\\\" ]; then echo 'MacWord2008 %s '\" & quoted form of POSIX path of (path to current application) & \" > \\\"$PIPE\\\"; else exit 1; fi;\"\n"
		 "on error\n"
		 "display alert \"Word could not communicate with Zotero. Please ensure that Firefox is open and try again.\" as critical\n"
						   "end try\n", W2008_SCRIPT_COMMANDS[i]]
						   dataUsingEncoding:NSUTF8StringEncoding]];
		[inFile closeFile];
		[outFile readDataToEndOfFile];
	}
	return STATUS_OK;
}

// Gets a string for an authorization failure
statusCode parseAuthorizationStatus(OSStatus status, const char file[],
									const char function[], unsigned int line) {
	if(status == errAuthorizationCanceled) {
		return STATUS_EXCEPTION_ALREADY_DISPLAYED;
	}
	
	NSString* err = [NSError errorWithDomain:NSOSStatusErrorDomain code:status
									userInfo:nil];
	throwError(err, file, function, line);
	return STATUS_EXCEPTION;
}

AuthorizationRef authorizationRef = NULL;
statusCode getAuthorizationRef(NSString* templatePath) {
	if(!authorizationRef) {
		AuthorizationItem authRightItems[] = {
			{kAuthorizationRightExecute, strlen(PATH_MKDIR), PATH_MKDIR, 0},
			{kAuthorizationRightExecute, strlen(PATH_RM), PATH_RM, 0},
			{kAuthorizationRightExecute, strlen(PATH_DITTO), PATH_DITTO, 0}
		};
		AuthorizationItem authEnvironmentItems[] = {
			{kAuthorizationEnvironmentPrompt, strlen(AUTH_PROMPT), AUTH_PROMPT,
				0},
			{kAuthorizationEnvironmentIcon, 0, NULL, 0}
		};
		AuthorizationRights authRights = {3, authRightItems};
		AuthorizationEnvironment authEnvironment = {2, authEnvironmentItems};
		
		//  Get path to Zotero icon from templatePath
		NSString* iconPathNS = [[templatePath stringByDeletingLastPathComponent]
								stringByAppendingPathComponent:@"zotero.png"];
		char* iconPath = copyNSString(iconPathNS);
		authEnvironmentItems[1].valueLength = strlen(iconPath);
		authEnvironmentItems[1].value = iconPath;
		
		OSStatus status;
		status = AuthorizationCreate(&authRights, &authEnvironment,
									 kAuthorizationFlagExtendRights | 
									 kAuthorizationFlagInteractionAllowed,
									 &authorizationRef);
		free(iconPath);
		if(status) {
			return parseAuthorizationStatus(status, __FILE__, __FUNCTION__,
											__LINE__-2);
		}
	}
	return STATUS_OK;
}

// Performs an mkdir operation, requesting privileges if necessary and
// optionally emptying the directory if it exists
statusCode performAuthorizedMkdir(NSString* templatePath, NSString *path,
								  BOOL emptyIfExists) {
	NSFileManager *fm = [NSFileManager defaultManager];
	BOOL isDirectory;
	if([fm fileExistsAtPath:path isDirectory:&isDirectory]) {
		if(emptyIfExists) {
			// File exists but we were asked to empty it
			if(![fm removeItemAtPath:path error:NULL]) {
				statusCode status = getAuthorizationRef(templatePath);
				if(status) return status;
				
				char *argv[3];
				argv[0] = "-rf";
				argv[1] = copyNSString(path);
				argv[2] = NULL;
				OSStatus authStatus = AuthorizationExecuteWithPrivileges(authorizationRef,
																		 PATH_RM,
																		 kAuthorizationFlagDefaults,
																		 argv, NULL);
				free(argv[1]);
				if(authStatus) {
					return parseAuthorizationStatus(authStatus, __FILE__,
													__FUNCTION__, __LINE__);
				}
			}
		} else {
			// Exists and we were not asked to empty
			if(!isDirectory) {
				NSString *err = [NSString stringWithFormat:
								@"Expected directory at %@", path];
				DIE(err);
			}
			return STATUS_OK;
		}
	}
	
	if(![fm createDirectoryAtPath:path
	  withIntermediateDirectories:YES attributes:nil error:NULL]) {
		statusCode status = getAuthorizationRef(templatePath);
		if(status) return status;

		char *argv[5];
		argv[0] = "-p";
		argv[1] = "-m";
		argv[2] = "775";
		argv[3] = copyNSString(path);
		argv[4] = NULL;
		OSStatus authStatus = AuthorizationExecuteWithPrivileges(authorizationRef,
																 PATH_MKDIR,
																 kAuthorizationFlagDefaults,
																 argv, NULL);
		free(argv[3]);
		if(authStatus) {
			return parseAuthorizationStatus(authStatus, __FILE__,
											__FUNCTION__, __LINE__);
		}
	}
	
	return STATUS_OK;
}

// Performs a copy using ditto, requesting privileges if necessary
statusCode performAuthorizedCopy(NSString* templatePath, NSString* path1,
								 NSString* path2) {
	// First try copying with copyfile (because it will happily clobber the
	// existing file, even if we don't have permissions on the directory,
	// whereas Cocoa APIs won't)
	if(!copyfile([path1 UTF8String], [path2 UTF8String], NULL, COPYFILE_ALL)) {
		return STATUS_OK;
	}
	
	// Create authorization reference
	statusCode status = getAuthorizationRef(templatePath);
	if(status) return status;
	
	// Run the tool using the authorization reference
	char *argv[3];
	argv[0] = copyNSString(path1);
	argv[1] = copyNSString(path2);
	argv[2] = NULL;
	OSStatus authStatus = AuthorizationExecuteWithPrivileges(authorizationRef,
															 PATH_DITTO,
															 kAuthorizationFlagDefaults,
															 argv, NULL);
	free(argv[0]);
	free(argv[1]);
	if(authStatus) {
		return parseAuthorizationStatus(authStatus, __FILE__,
										__FUNCTION__, __LINE__);
	}
	
	return STATUS_OK;
}