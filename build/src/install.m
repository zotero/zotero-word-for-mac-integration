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

#define AUTH_PROMPT "Zotero requires administrator "\
	"privileges to complete installation of the Word plugin."
#define PATH_MKDIR "/bin/mkdir"
#define PATH_RM "/bin/rm"
#define PATH_DITTO "/usr/bin/ditto"
#define PATH_CHOWN "/usr/sbin/chown"
#define PATH_CHMOD "/bin/chmod"

statusCode setTemplateTypeCreator(NSString* templatePath, bool);
statusCode installContainerTemplate(NSString*);
statusCode installAppleScript(NSString*);
statusCode parseAuthorizationStatus(OSStatus status, const char file[],
									const char function[], unsigned int line);
statusCode performAuthorizedAction(NSString* templatePath,
								   const char* commandPath, char* const* argv);
statusCode performAuthorizedMkdir(NSString* templatePath, NSString *path,
								  BOOL emptyIfExists);
statusCode performAuthorizedCopy(NSString* templatePath, NSString* path1,
								 NSString* path2);
@interface PipeReader : NSObject
- (void) getData:(NSNumber*)progressValue;
@end

// Installs all scripts and templates, given the path to Zotero.dot
statusCode install(const char zoteroDotPath[], const char zoteroDotmPath[], const char zoteroScptPath[]) {
	HANDLE_EXCEPTIONS_BEGIN
	NSFileManager *fm = [NSFileManager defaultManager];
	NSString* dotPathNS = [NSString stringWithUTF8String:zoteroDotPath];
	NSString* dotmPathNS = [NSString stringWithUTF8String:zoteroDotmPath];
	NSString* scptPathNS = [NSString stringWithUTF8String:zoteroScptPath];
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
	
	BOOL shouldInstallContainerTemplate = NO, installed = NO, oldWordFound = NO, shouldPromptAboutWordUpdate = NO;
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
		
		if (!version) continue;
		NSUInteger dotIndex = [version rangeOfString:@"."].location;
		if (dotIndex == NSNotFound) continue;
		NSUInteger secondDotIndex = [version rangeOfString:@"." options:NSLiteralSearch range:NSMakeRange(dotIndex, [version length] - dotIndex)].location;
		
		NSInteger majorVersion = [[version substringToIndex:dotIndex] intValue];
		NSInteger minorVersion = 0;
		if (secondDotIndex != NSNotFound) {
			 minorVersion = [[version substringWithRange:NSMakeRange(dotIndex+1, secondDotIndex)] intValue];
		}
		
		// Prompt that Zotero plugin doesn't work on versions older than 2016
		oldWordFound = oldWordFound || majorVersion < 15;
		
        // Install template into container directory for Word 2016
		shouldInstallContainerTemplate = shouldInstallContainerTemplate || majorVersion >= 15;
		
		// Prompt pre-16 users to updates
		shouldPromptAboutWordUpdate = shouldPromptAboutWordUpdate || (majorVersion == 15);
	}
	
    if (shouldInstallContainerTemplate) {
		ENSURE_OK(installContainerTemplate(dotmPathNS))
		ENSURE_OK(installAppleScript(scptPathNS))
        installed = true;
    }
	
	if(!installed) {
		if (oldWordFound) {
			DIE(@"The Zotero Word plugin "
				"requires Word 2016 or later. Please upgrade Word.");
		} else {
			DIE(@"Word does not appear to be installed on this computer. "
				"Please install Microsoft Word and try again.");
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
		NSArray* apps = [workspace runningApplications];
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
		dispatch_sync(dispatch_get_main_queue(), ^{
			NSAlert *alert = [NSAlert alertWithMessageText:@"The Zotero Word "
							  "plugin has been successfully installed, but "
							  "Word must be restarted before it can be used."
											 defaultButton:nil
										   alternateButton:nil
											   otherButton:nil
								 informativeTextWithFormat:@"Please restart Word "
							  "before continuing."];
			[alert runModal];
		});
	}
	if (shouldPromptAboutWordUpdate) {
		dispatch_sync(dispatch_get_main_queue(), ^{
			NSAlert *alert = [NSAlert alertWithMessageText:@"The Zotero Word "
							  "plugin has been successfully installed, but "
							  "Word 2016 must be updated before the plugin can be used."
											 defaultButton:nil
										   alternateButton:nil
											   otherButton:nil
								 informativeTextWithFormat:@"Please update Word 2016 to "
							  "version 16.16.27 or later."];
			[alert runModal];
		});
	}
	
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Sets the type and creator code on the template
statusCode setTemplateTypeCreator(NSString* templatePath, bool dieOnError) {
    NSError *err = nil;
    if(![[NSFileManager defaultManager] setAttributes:[NSDictionary dictionaryWithObjectsAndKeys:
                           [NSNumber numberWithUnsignedInt:'MSWD'],
                           NSFileHFSCreatorCode,
                           [NSNumber numberWithUnsignedInt:'W8TN'],
                           NSFileHFSTypeCode, nil]
             ofItemAtPath:templatePath error:&err]) {
		if (dieOnError) {
			DIE([err localizedDescription]);
		}
    }

    return STATUS_OK;
}

statusCode installContainerTemplate(NSString* dotmPathNS) {
    NSString *startupDirectory = [NSHomeDirectory() stringByAppendingPathComponent:
								  @"Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word"];
	NSString *newTemplatePath = [startupDirectory stringByAppendingPathComponent:
								 @"Zotero.dotm"];
	NSString *oldTemplatePath = [startupDirectory stringByAppendingPathComponent:
								 @"Zotero.dot"];
	NSFileManager *fm = [NSFileManager defaultManager];
	
	NSError *err = nil;
	if(![fm fileExistsAtPath:startupDirectory]) {
		// Word startup directory does not exist yet. Warn user they will have to
		// restart Firefox/Zotero
		if(![fm createDirectoryAtPath:startupDirectory withIntermediateDirectories:YES
		     attributes:nil error:&err]) {
			DIE([err localizedDescription]);
		}

		dispatch_sync(dispatch_get_main_queue(), ^{
			NSAlert *alert = [NSAlert alertWithMessageText:@"The Zotero Word "
							  "plugin has been successfully installed, but "
							  "Zotero must be restarted before it can be used."
											 defaultButton:nil
										   alternateButton:nil
											   otherButton:nil
								 informativeTextWithFormat:@"Please restart Zotero "
							  "before continuing."];
			[alert runModal];
		});
	}
	if([fm fileExistsAtPath:newTemplatePath]) {
		if(![fm removeItemAtPath:newTemplatePath error:&err]) {
			DIE([err localizedDescription]);
		}
	}
	// Remove old template path
	if([fm fileExistsAtPath:oldTemplatePath]) {
		if(![fm removeItemAtPath:oldTemplatePath error:&err]) {
			DIE([err localizedDescription]);
		}
	}
	if(![fm copyItemAtPath:dotmPathNS toPath:newTemplatePath error:&err]) {
		DIE([err localizedDescription]);
    }
    ENSURE_OK(setTemplateTypeCreator(newTemplatePath, true))
	
	return STATUS_OK;
}

statusCode installAppleScript(NSString* scriptPathNS) {
	// Get path to startup folder
	NSString *applicationScriptsDirectory = [NSHomeDirectory() stringByAppendingPathComponent:
								  @"Library/Application Scripts/com.microsoft.Word"];
	
	// Make sure directory exists, but don't empty it
	statusCode status = performAuthorizedMkdir(scriptPathNS, applicationScriptsDirectory,
											   false);
	if(status) return status;

	NSString *newScriptPath = [applicationScriptsDirectory
								 stringByAppendingPathComponent:
								 @"Zotero.scpt"];
	
	status = performAuthorizedCopy(scriptPathNS, scriptPathNS, newScriptPath);
	if(status) return status;
	
	// Fix template type and creator
	status = setTemplateTypeCreator(newScriptPath, false);
	if(status) return status;
	
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
