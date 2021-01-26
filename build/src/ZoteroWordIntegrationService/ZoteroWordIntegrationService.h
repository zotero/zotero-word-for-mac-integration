//
//  ZoteroWordIntegrationService.h
//  ZoteroWordIntegrationService
//
//  Created by Adomas Venckauskas on 2021-01-06.
//

#import <Foundation/Foundation.h>
#import "../ZoteroWordIntegrationServiceProtocol.h"
#import "../zoteroMacWordIntegration.h"

// This object implements the protocol which we have defined. It provides the actual behavior for the service. It is 'exported' by the service to make it available to the process hosting the service over an NSXPCConnection.
@interface ZoteroWordIntegrationService : NSObject <ZoteroWordIntegrationServiceProtocol>

@property document_t *doc;

@end
