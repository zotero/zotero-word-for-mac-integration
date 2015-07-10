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
#define PATH_CHOWN "/usr/sbin/chown"
#define PATH_CHMOD "/bin/chmod"

statusCode installTemplateIntoStartupDirectory(NSString* templatePath, NSString* path);
statusCode setTemplateTypeCreator(NSString* templatePath);
statusCode installScripts(NSString* templatePath);
statusCode parseAuthorizationStatus(OSStatus status, const char file[],
									const char function[], unsigned int line);
statusCode performAuthorizedAction(NSString* templatePath,
								   const char* commandPath, char* const* argv);
statusCode performAuthorizedMkdir(NSString* templatePath, NSString *path,
								  BOOL emptyIfExists);
statusCode performAuthorizedCopy(NSString* templatePath, NSString* path1,
								 NSString* path2);
statusCode getScriptItemsDirectories(NSMutableArray* scriptFolders);
statusCode writeScriptNS(NSString* scriptPath, NSString* scriptContent);
@interface PipeReader : NSObject
- (void) getData:(NSNumber*)progressValue;
@end

// Installs all scripts and templates, given the path to Zotero.dot
statusCode install(const char templatePath[]) {
	HANDLE_EXCEPTIONS_BEGIN
	NSFileManager *fm = [NSFileManager defaultManager];
	NSString* templatePathNS = [NSString stringWithUTF8String:templatePath];
	FinderApplication* finder = nil;
	
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
    [wordLocations addObject:[applicationsFolder stringByAppendingPathComponent:
                              @"Microsoft Word.app"]];
	
	// Look for Word bundle
	NSWorkspace *workspace = [NSWorkspace sharedWorkspace];
	NSString* testPath = [workspace absolutePathForAppBundleWithIdentifier:
						  @"com.Microsoft.Word"];
	if(testPath) [wordLocations addObject:testPath];
	
	// Look for Word by creator code
	CFURLRef testURL;
	if(!LSFindApplicationForInfo('MSWD', NULL, NULL, NULL, &testURL) && testURL) {
		[wordLocations addObject:[(NSURL*)testURL path]];
	}
	
	BOOL shouldInstallScripts = NO, shouldInstallContainerTemplate = NO, installed = NO, wordXFound = NO;
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
		
		// Install scripts for Word 2008 and 2011
		shouldInstallScripts = shouldInstallScripts || (intVersion >= 12 && intVersion < 15);
        if(intVersion == 11 || intVersion == 14) {
            // Install template into startup directory for Word 2004 or Word 2011
			statusCode status = installTemplateIntoStartupDirectory(templatePathNS, path);
			if(status) return status;
			installed = true;
        } else if(intVersion == 15) {
            // Install template into container directory for Word 2016
            shouldInstallContainerTemplate = true;
        }
	}
	
	if(shouldInstallScripts) {
		statusCode status = installScripts(templatePathNS);
		if(status) return status;
		installed = true;
	}

    if (shouldInstallContainerTemplate) {
        NSString *newTemplatePath = [NSHomeDirectory() stringByAppendingPathComponent:
                                     @"Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word/Zotero.dot"];
        NSError *err = nil;
        if([fm fileExistsAtPath:newTemplatePath]) {
            if(![fm removeItemAtPath:newTemplatePath error:&err]) {
                DIE([err localizedDescription]);
            }
        }
        if(![fm copyItemAtPath:templatePathNS toPath:newTemplatePath error:&err]) {
            DIE([err localizedDescription]);
        }
        statusCode status = setTemplateTypeCreator(newTemplatePath);
        if(status) return status;
        installed = true;
    }
	
	if(!installed) {
		if(wordXFound) {
			DIE(@"While an installation of Word X was found, Zotero Word for "
				"Mac Integration requires Word 2004 or later. Please upgrade "
				"Word to use Zotero Word Integration.");
		} else {
			DIE(@"Neither Word does not appear to be installed on "
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
	HANDLE_EXCEPTIONS_END
}

// Sets the type and creator code on the template
statusCode setTemplateTypeCreator(NSString* templatePath) {
    NSError *err = nil;
    if(![[NSFileManager defaultManager] setAttributes:[NSDictionary dictionaryWithObjectsAndKeys:
                           [NSNumber numberWithUnsignedInt:'MSWD'],
                           NSFileHFSCreatorCode,
                           [NSNumber numberWithUnsignedInt:'W8TN'],
                           NSFileHFSTypeCode, nil]
             ofItemAtPath:templatePath error:&err]) {
        DIE([err localizedDescription]);
    }

    return STATUS_OK;
}

// Installs a template to a given path
statusCode installTemplateIntoStartupDirectory(NSString* templatePath, NSString* path) {
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
	NSString *newTemplatePath = [startupDirectory
								stringByAppendingPathComponent:
								 @"Zotero.dot"];
	status = performAuthorizedCopy(templatePath, templatePath, newTemplatePath);
	if(status) return status;
	
	// Fix template type and creator
    status = setTemplateTypeCreator(newTemplatePath);
    if(status) return status;
	
	return STATUS_OK;
}

// Gets the location of the Script Menu Items directory
statusCode getScriptItemsDirectories(NSMutableArray* scriptFolders) {
	NSFileManager *fm = [NSFileManager defaultManager];
	
	// Look for Script Menu Items folder in
	// Application Support/Microsoft/Office/Word Script Menu Items
	NSArray *temp = NSSearchPathForDirectoriesInDomains(NSApplicationSupportDirectory,
														NSUserDomainMask, YES);
	if(!temp || ![temp count]) DIE(@"Could not find Application Support directory");
	NSString *testPath = [[temp objectAtIndex:0] stringByAppendingPathComponent:
						   @"Microsoft/Office/Word Script Menu Items"];
	BOOL isDirectory;
	if([fm fileExistsAtPath:testPath isDirectory:&isDirectory]
	   && isDirectory) {
		[scriptFolders addObject:testPath];
	}
	
	// Look for Script Menu Items folder in Microsoft User Data folder
	temp = NSSearchPathForDirectoriesInDomains(NSDocumentDirectory,
												NSUserDomainMask, YES);
	if(!temp || ![temp count]) DIE(@"Could not find Documents directory");
	NSString* documentsFolder = [temp objectAtIndex:0];
	
	temp = NSSearchPathForDirectoriesInDomains(NSLibraryDirectory,
											   NSUserDomainMask, YES);
	if(!temp || ![temp count]) DIE(@"Could not find Library directory");
	NSString* preferencesFolder = [[temp objectAtIndex:0]
								   stringByAppendingPathComponent:
								   @"Preferences"];
	
	// Look for Script Menu Items folder in various other places
	NSString* potentialMicrosoftUserDataNames[] = {@"Microsoft User Data",
		@"Microsoft-Benutzerdaten", @"Données Utilisateurs Microsoft",
		@"Datos del Usuario de Microsoft", @"Datos de Usuario de Microsoft",
		@"Dati utente Microsoft", NULL};
	for(unsigned short i=0; potentialMicrosoftUserDataNames[i] != NULL;
		i++) {
		NSString *testPath = [[documentsFolder stringByAppendingPathComponent:
							   potentialMicrosoftUserDataNames[i]]
							  stringByAppendingPathComponent:
							  @"Word Script Menu Items"];
		BOOL isDirectory;
		if([fm fileExistsAtPath:testPath isDirectory:&isDirectory]
		   && isDirectory) {
			[scriptFolders addObject:testPath];
			break;
		} else {
			testPath = [[preferencesFolder stringByAppendingPathComponent:
						 potentialMicrosoftUserDataNames[i]]
						stringByAppendingPathComponent:
						@"Word Script Menu Items"];
			if([fm fileExistsAtPath:testPath isDirectory:&isDirectory]
			   && isDirectory) {
				[scriptFolders addObject:testPath];
				break;
			}
		}
	}
	
	// We couldn't find Script Menu Items folder, so ask user to locate it
	if(![scriptFolders count]) {
		NSOpenPanel* openPanel = [NSOpenPanel openPanel];
		[openPanel setCanChooseFiles:NO];
		[openPanel setCanChooseDirectories:YES];
		[openPanel setAllowsMultipleSelection:NO];
		[openPanel setMessage:@"Select the Word Script Menu Items folder, "
		 "usually located in Documents/Microsoft User Data"];
		[openPanel setDirectory:documentsFolder];
		if([openPanel runModal] == NSFileHandlingPanelOKButton) {
			[scriptFolders addObject:[[[openPanel URLs] objectAtIndex:0] path]];
		} else {
			return STATUS_EXCEPTION_ALREADY_DISPLAYED;
		}
	}
	
	return STATUS_OK;
}

statusCode getScriptItemsDirectory(char** scriptFolder) {
	HANDLE_EXCEPTIONS_BEGIN
	NSMutableArray* scriptFolders = [NSMutableArray array];
	ENSURE_OK(getScriptItemsDirectories(scriptFolders))
	*scriptFolder = copyNSString([scriptFolders objectAtIndex:0]);
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

statusCode writeScriptNS(NSString* scriptPath, NSString* scriptContent) {
	NSTask *task;
	task = [[NSTask alloc] init];
	[task autorelease];
	[task setLaunchPath: @"/usr/bin/osacompile"];
	NSPipe *inPipe = [NSPipe pipe];
	[task setStandardInput:inPipe];
	NSFileHandle *inFile = [inPipe fileHandleForWriting];
	
	SInt32 versMaj, versMin;
	Gestalt(gestaltSystemVersionMajor, &versMaj);
	Gestalt(gestaltSystemVersionMinor, &versMin);
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
	[inFile writeData:[scriptContent dataUsingEncoding:NSUTF8StringEncoding]];
	[inFile closeFile];
	
	return STATUS_OK;
}

statusCode writeScript(char* scriptPath, char* scriptContent) {
	HANDLE_EXCEPTIONS_BEGIN
	return writeScriptNS([NSString stringWithUTF8String:scriptPath],
						 [NSString stringWithUTF8String:scriptContent]);
	HANDLE_EXCEPTIONS_END
}

statusCode installScripts(NSString* templatePath) {
	NSMutableArray* scriptFolders = [NSMutableArray array];
	ENSURE_OK(getScriptItemsDirectories(scriptFolders))
	
	for(NSString *scriptFolder in scriptFolders) {
		// Clobber the Zotero folder
		NSString* zoteroFolder = [scriptFolder
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
			NSString* scriptPath = [zoteroFolder stringByAppendingPathComponent:
									[NSString stringWithUTF8String: 
									 W2008_SCRIPT_NAMES[i]]];
			NSString* scriptContent = [NSString stringWithFormat:@"try\n"
									   "do shell script \"PIPE=\\\"/Users/Shared/.zoteroIntegrationPipe_$LOGNAME\\\";  if [ ! -e \\\"$PIPE\\\" ]; then PIPE=~/.zoteroIntegrationPipe; fi; if [ -e \\\"$PIPE\\\" ]; then echo 'MacWord2008 %s '\" & quoted form of POSIX path of (path to current application) & \" > \\\"$PIPE\\\"; else exit 1; fi;\"\n"
									   "on error\n"
									   "display alert \"Word could not communicate with Zotero. Please ensure that Zotero Standalone or Firefox is open and try again.\" as critical\n"
									   "end try\n", W2008_SCRIPT_COMMANDS[i]];
			ENSURE_OK(writeScriptNS(scriptPath, scriptContent))
		}
	}
	return STATUS_OK;
}

AuthorizationRef authorizationRef = NULL;
// Get authorization reference, perform a command as root, and block until it is
// complete
statusCode performAuthorizedAction(NSString* templatePath,
								   const char* commandPath, char* const* argv) {
	if(!authorizationRef) {
		AuthorizationItem authRightItems[] = {
			{kAuthorizationRightExecute, strlen(PATH_MKDIR), PATH_MKDIR, 0},
			{kAuthorizationRightExecute, strlen(PATH_RM), PATH_RM, 0},
			{kAuthorizationRightExecute, strlen(PATH_DITTO), PATH_DITTO, 0},
			{kAuthorizationRightExecute, strlen(PATH_CHOWN), PATH_CHOWN, 0},
			{kAuthorizationRightExecute, strlen(PATH_CHMOD), PATH_CHMOD, 0}
		};
		AuthorizationItem authEnvironmentItems[] = {
			{kAuthorizationEnvironmentPrompt, strlen(AUTH_PROMPT), AUTH_PROMPT,
				0},
			{kAuthorizationEnvironmentIcon, 0, NULL, 0}
		};
		AuthorizationRights authRights = {5, authRightItems};
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
		CHECK_OSSTATUS(status)
	}
	
	// Execute command
	FILE *communicationsPipe;
	OSStatus authStatus = AuthorizationExecuteWithPrivileges(authorizationRef,
															 commandPath,
															 kAuthorizationFlagDefaults,
															 argv, &communicationsPipe);
	CHECK_OSSTATUS(authStatus)
	
	// Block until command closed stdout
	char buffer[1024];
	while(fread(&buffer, 1024, 1, communicationsPipe)) {};
	fclose(communicationsPipe);
	
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
				char *argv[3];
				argv[0] = "-rf";
				argv[1] = copyNSString(path);
				argv[2] = NULL;
				ENSURE_OK(performAuthorizedAction(templatePath, PATH_RM, argv))
				free(argv[1]);
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
		char *argv[5];
		argv[0] = "-p";
		argv[1] = "-m";
		argv[2] = "775";
		argv[3] = copyNSString(path);
		argv[4] = NULL;
		ENSURE_OK(performAuthorizedAction(templatePath, PATH_MKDIR, argv))
		free(argv[3]);
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
	
	// Set up arguments
	char *argv[3];
	argv[0] = copyNSString(path1);
	argv[1] = copyNSString(path2);
	argv[2] = NULL;
	ENSURE_OK(performAuthorizedAction(templatePath, PATH_DITTO, argv))
	free(argv[0]);
	free(argv[1]);
	
	// Make sure we own the template
	char uid[11];
	if(snprintf(uid, 11, "%d", getuid()) >= 11) {
		DIE(@"UID too long")
	}
	argv[0] = uid;
	argv[1] = copyNSString(path2);
	argv[2] = NULL;
	ENSURE_OK(performAuthorizedAction(templatePath, PATH_CHOWN, argv))
	free(argv[1]);
	
	// Make sure we have read access on the parent
	argv[0] = "755";
	argv[1] = copyNSString([path2 stringByDeletingLastPathComponent]);
	argv[2] = NULL;
	ENSURE_OK(performAuthorizedAction(templatePath, PATH_CHMOD, argv))
	free(argv[1]);
	
	return STATUS_OK;
}