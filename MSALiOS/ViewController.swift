//------------------------------------------------------------------------------
//
// Copyright (c) Microsoft Corporation.
// All rights reserved.
//
// This code is licensed under the MIT License.
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files(the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and / or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions :
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//
//------------------------------------------------------------------------------

import UIKit
import MSAL

/// 😃 A View Controller that will respond to the events of the Storyboard.

class ViewController: UIViewController, UITextFieldDelegate, URLSessionDelegate {
    
    // Update the below to your client ID you received in the portal. The below is for running the demo only
    
    //TODO: come from info.plist
    let kClientID = "d2a20985-1b2c-4a50-8116-faa8b508fa2a"
    
    // These settings you don't need to edit unless you wish to attempt deeper scenarios with the app.
    let kIssuer = "https://login.microsoftonline.com/brandwedir.onmicrosoft.com/v2.0"
    let kGraphURI = "https://graph.microsoft.com/v1.0/me/"
    let kScopes: [String] = ["https://graph.microsoft.com/user.read https://graph.microsoft.com/Files.ReadWrite.AppFolder"]
    let kAuthority = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
    
    var accessToken = String()
    var applicationContext = MSALPublicClientApplication.init()
    let oneDriveManager: OneDriveManager = OneDriveManager()
    
    var appFolderId: String?  // This is not static by design
    
    @IBOutlet weak var loggingText: UITextView!
    @IBOutlet weak var signoutButton: UIButton!
    
    /**
     This button will invoke the authorization flow.
     */
    
    @IBAction func callGraphButton(_ sender: UIButton) {
        do {
            
            // We check to see if we have a current logged in user. If we don't, then we need to sign someone in.
            // We throw an interactionRequired so that we trigger the interactive signin.
            
            
            if  try self.applicationContext.users().isEmpty {
                throw NSError.init(domain: "MSALErrorDomain", code: MSALErrorCode.interactionRequired.rawValue, userInfo: nil)
            } else {
                
                /**
                 
                 Acquire a token for an existing user silently
                 
                 - forScopes: Permissions you want included in the access token received
                 in the result in the completionBlock. Not all scopes are
                 gauranteed to be included in the access token returned.
                 - User: A user object that we retrieved from the application object before that the
                 authentication flow will be locked down to.
                 - completionBlock: The completion block that will be called when the authentication
                 flow completes, or encounters an error.
                 */
                
                try self.applicationContext.acquireTokenSilent(forScopes: self.kScopes, user: applicationContext.users().first) { (result, error) in
                    
                    if error == nil {
                        self.accessToken = (result?.accessToken)!
                        self.oneDriveManager.accessToken = self.accessToken
                        
                        DispatchQueue.main.async {
                            self.loggingText.text = "Refreshing token silently)"
                            self.loggingText.text = "Refreshed Access token is \(self.accessToken)"
                            
                            self.signoutButton.isEnabled = true;
                        }
                        self.getContentWithToken()
                        self.getAppMetadata()
                        
                        self.oneDriveManager.getAppFolderId(completion: { (result: OneDriveManagerResult, appFolderId) -> Void in
                            switch(result) {
                            case .Success:
                                self.appFolderId = appFolderId
                                print ("\(self.appFolderId))")
                                
                                /**
                                 let fileName = "test.txt"
                                 
                                 self.oneDriveManager.createTextFile(fileName: fileName, folderId: appFolderId!, completion: { (result: OneDriveManagerResult, webUrl)  -> Void in
                                 switch(result) {
                                 case .Success:
                                 print ("success (\(webUrl)")
                                 case .Failure(let error):
                                 print("\(error)")
                                 }
                                 })
                                 
                                 let folderName = "testing"
                                 self.oneDriveManager.createFolder(folderName: folderName, folderId: appFolderId!, completion: { (result: OneDriveManagerResult)  -> Void in
                                 switch(result) {
                                 case .Success:
                                 print ("success")
                                 case .Failure(let error):
                                 print("\(error)")
                                 }
                                 })
                                 
                                 **/
                                
                                self.oneDriveManager.createUploadSession(fileName: "rose.jpg", folderId: appFolderId!,  completion: { (result: OneDriveManagerResult, uploadUrl, expirationDateTime, nextExpectedRanges) -> Void in
                                    switch(result) {
                                    case .Success:
                                        print ("success on creating session (\(String(describing: uploadUrl)) (\(String(describing: expirationDateTime))")
                                        
                                        self.oneDriveManager.uploadBytes(uploadUrl: uploadUrl!,  completion: { (result: OneDriveManagerResult, webUrl, fileId) -> Void in
                                            switch(result) {
                                            case .Success:
                                                print ("Web Url of file \(webUrl)")
                                                print ("FileId of file \(fileId)")
                                                
                                                self.oneDriveManager.createSharingLink(fileId: fileId!, completion: { (result: OneDriveManagerResult, sharingUrl) -> Void in
                                                    switch(result) {
                                                    case .Success:
                                                        print ("Sharing Url of file \(sharingUrl)")
                                                        
                                                    case .Failure(let error):
                                                        print("\(error)")
                                                    }
                                                })
                                                
                                            case .Failure(let error):
                                                print("\(error)")
                                            }
                                        })
                                    case .Failure(let error):
                                        print("\(error)")
                                    }
                                })
                                
                            case .Failure(let error):
                                print("\(error)")
                            }
                        })
                    } else {
                        DispatchQueue.main.async {
                            self.loggingText.text = "Could not acquire token silently: \(error ?? "No error informarion" as! Error)"
                        }
                    }
                }
            }
        }  catch let error as NSError {
            
            // interactionRequired means we need to ask the user to sign-in. This usually happens
            // when the user's Refresh Token is expired or if the user has changed their password
            // among other possible reasons.
            
            if error.code == MSALErrorCode.interactionRequired.rawValue {
                
                self.applicationContext.acquireToken(forScopes: self.kScopes) { (result, error) in
                    if error == nil {
                        self.accessToken = (result?.accessToken)!
                        self.loggingText.text = "Access token is \(self.accessToken)"
                        self.signoutButton.isEnabled = true;
                        self.getContentWithToken()
                        
                    } else  {
                        self.loggingText.text = "Could not acquire token: \(error ?? "No error informarion" as! Error)"
                    }
                }
                
            }
            
        } catch {
            
            // This is the catch all error.
            
            self.loggingText.text = "Unable to acquire token. Got error: \(error)"
            
        }
    }
    
    
    /**
     This button will invoke the call to the Microsoft Graph API. It uses the
     built in Swift libraries to create a connection.
     Pay attention to the error case below. It shows you how to
     detect a `UserInteractionRequired` Error case and present the `acquireToken()`
     method again for the user to sign in. This usually happens if
     the Refresh Token has expired or the user has changed their
     password.
     
     */
    
    func getContentWithToken() {
        
        let sessionConfig = URLSessionConfiguration.default
        
        let url = URL(string: kGraphURI)
        var request = URLRequest(url: url!)
        request.setValue("Bearer \(self.accessToken)", forHTTPHeaderField: "Authorization")
        let urlSession = URLSession(configuration: sessionConfig, delegate: self, delegateQueue: OperationQueue.main)
        
        urlSession.dataTask(with: request) { data, response, error in
            
            let result = try? JSONSerialization.jsonObject(with: data!, options: [])
            if result != nil {
                
                self.loggingText.text = result.debugDescription
            }
            }.resume()
    }
    
    func getAppMetadata() {
        let sessionConfig = URLSessionConfiguration.default
        let url = URL(string: "https://graph.microsoft.com/v1.0/me/drive/special/approot")
        var request = URLRequest(url: url!)
        request.setValue("Bearer \(self.accessToken)", forHTTPHeaderField: "Authorization")
        let urlSession = URLSession(configuration: sessionConfig, delegate: self, delegateQueue: OperationQueue.main)
        
        urlSession.dataTask(with: request) { data, response, error in
            
            let result = try? JSONSerialization.jsonObject(with: data!, options: [])
            if result != nil {
                
                self.loggingText.text = result.debugDescription
            }
            }.resume()
    }
    
    /**
     This button will invoke the signout APIs to clear the token cache.
     
     */
    
    @IBAction func signoutButton(_ sender: UIButton) {
        
        do {
            
            /**
             Removes all tokens from the cache for this application for the provided user
             
             - user:    The user to remove from the cache
             */
            
            try self.applicationContext.remove(self.applicationContext.users().first)
            self.signoutButton.isEnabled = false;
            
        } catch let error {
            self.loggingText.text = "Received error signing user out: \(error)"
        }
    }
    
    override func viewDidLoad() {
        super.viewDidLoad()
        
        do {
            
            /**
             
             Initialize a MSALPublicClientApplication with a given clientID and authority
             
             - clientId:     The clientID of your application, you should get this from the app portal.
             - authority:    A URL indicating a directory that MSAL can use to obtain tokens. In Azure AD
             it is of the form https://<instance/<tenant>, where <instance> is the
             directory host (e.g. https://login.microsoftonline.com) and <tenant> is a
             identifier within the directory itself (e.g. a domain associated to the
             tenant, such as contoso.onmicrosoft.com, or the GUID representing the
             TenantID property of the directory)
             - Parameter error       The error that occurred creating the application object, if any, if you're
             not interested in the specific error pass in nil.
             
             */
            
            self.applicationContext = try MSALPublicClientApplication.init(clientId: kClientID, authority: kAuthority)
            
        } catch {
            
            self.loggingText.text = "Unable to create Application Context"
            
        }
    }
    
    override func didReceiveMemoryWarning() {
        super.didReceiveMemoryWarning()
        // Dispose of any resources that can be recreated.
    }
    
    override func viewWillAppear(_ animated: Bool) {
        
        if self.accessToken.isEmpty {
            
            signoutButton.isEnabled = false;
            
        }
    }
    
}

