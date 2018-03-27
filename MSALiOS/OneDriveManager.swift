/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import UIKit

enum OneDriveManagerResult {
    case Success
    case Failure(OneDriveAPIError)
}

enum OneDriveAPIError: Error {
    case ResourceNotFound
    case JSONParseError
    case UnspecifiedError(URLResponse?)
    case GeneralError(Error?)
}

class OneDriveManager : NSObject {
    //TODO: hard coded end point
    var baseURL: String = "https://graph.microsoft.com/v1.0/"

    var accessToken = String()

    override init() {
        super.init()
    }
    
    
    // MARK: Step 1 - folder creation/retrieval
    func getAppFolderId(completion: @escaping (OneDriveManagerResult, _ appFolderId: String?) -> Void) {
        let request = NSMutableURLRequest(url: URL(string: "\(baseURL)/me/drive/special/approot:/")!)
        
        request.httpMethod = "GET"
        request.setValue("application/json, text/plain, */*", forHTTPHeaderField: "Accept")
        request.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")
        
        let task = URLSession.shared.dataTask(with: request as URLRequest, completionHandler: {
            (data, response, error) -> Void in
            
            if let someError = error {
                completion(OneDriveManagerResult.Failure(OneDriveAPIError.GeneralError(someError)), nil)
                return
            }
            
            let statusCode = (response as! HTTPURLResponse).statusCode
            print("status code = \(statusCode)")
            
            switch(statusCode) {
                case 200:
                    do{
                        let json = try JSONSerialization.jsonObject(with: data!, options: JSONSerialization.ReadingOptions()) as! [String:Any]
                        print((json.description)) // outputs whole JSON
                        
                        guard let folderId = json["id"] as? String else {
                            completion(OneDriveManagerResult.Failure(OneDriveAPIError.UnspecifiedError(response)), nil)
                            return
                        }
                        completion(OneDriveManagerResult.Success, folderId)
                    }
                    catch{
                        completion(OneDriveManagerResult.Failure(OneDriveAPIError.JSONParseError), nil)
                    }
                
                case 404:
                    completion(OneDriveManagerResult.Failure(OneDriveAPIError.ResourceNotFound), nil)
                
                default:
                    completion(OneDriveManagerResult.Failure(OneDriveAPIError.UnspecifiedError(response)), nil)
                
            }
        })
        
        task.resume()
    }
    
    func createTextFile(fileName:String, folderId:String, completion: @escaping (OneDriveManagerResult, _ appFolderId: String?) -> Void) {
        
        let config = URLSessionConfiguration.default
        let session = URLSession(configuration: config)
        
        let request = NSMutableURLRequest(url: NSURL(string: "\(baseURL)/me/drive/special/approot:/\(fileName):/content")! as URL)
        request.httpMethod = "PUT"
        request.setValue("text/plain", forHTTPHeaderField: "Content-Type")
        request.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")
        request.httpBody = ("This is a test text file" as NSString).data(using: String.Encoding.utf8.rawValue)
        
        let task = session.dataTask(with: request as URLRequest, completionHandler: {
            (data, response, error) -> Void in
            
            if let someError = error {
                completion(OneDriveManagerResult.Failure(OneDriveAPIError.GeneralError(someError)), nil)
                return
            }
            
            let statusCode = (response as! HTTPURLResponse).statusCode
            print("status code = \(statusCode)")
            
            switch(statusCode) {
                case 200, 201:
                    do {
                        let jsonResponse = try JSONSerialization.jsonObject(with: data!, options: [])  as? [String: Any]
                        print((jsonResponse?.description)!) // outputs whole JSON
                        
                        guard let webUrl = jsonResponse!["webUrl"] as? String else {
                            completion(OneDriveManagerResult.Failure(OneDriveAPIError.UnspecifiedError(response)), nil)
                            return
                        }
                        completion(OneDriveManagerResult.Success, webUrl)
                    }
                    catch{
                        completion(OneDriveManagerResult.Failure(OneDriveAPIError.JSONParseError), nil)
                    }
                default:
                    completion(OneDriveManagerResult.Failure(OneDriveAPIError.UnspecifiedError(response)),nil)
            }
        })
        task.resume()
    }
    
    func createUploadSession(fileName:String, folderId:String, completion: @escaping (OneDriveManagerResult, _ uploadUrl: String?, _ expirationDateTime: String?, _ nextExpectedRanges: [String]?) -> Void) {
        
        let config = URLSessionConfiguration.default
        let session = URLSession(configuration: config)
        
        let request = NSMutableURLRequest(url: NSURL(string: "\(baseURL)/me/drive/special/approot:/\(fileName):/createUploadSession")! as URL)
        request.httpMethod = "POST"
        request.setValue("application/json", forHTTPHeaderField: "Content-Type")
        request.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")
        let emptyParams = Dictionary<String, String>()
        let params = ["item": [
                      "@microsoft.graph.conflictBehavior":"rename",
                      "name":fileName]] as [String : Any]
        
        request.httpBody = try! JSONSerialization.data(withJSONObject: params, options: JSONSerialization.WritingOptions())
        
        let task = session.dataTask(with: request as URLRequest, completionHandler: {
            (data, response, error) -> Void in
            
            if let someError = error {
                completion(OneDriveManagerResult.Failure(OneDriveAPIError.GeneralError(someError)), nil, nil, nil)
                return
            }
            
            let statusCode = (response as! HTTPURLResponse).statusCode
            print("status code = \(statusCode)")
            
            switch(statusCode) {
            case 200, 201:
                do {
                    let jsonResponse = try JSONSerialization.jsonObject(with: data!, options: [])  as? [String: Any]
                    print((jsonResponse?.description)!) // outputs whole JSON
                    
                    guard let uploadUrl = jsonResponse!["uploadUrl"] as? String else {
                        completion(OneDriveManagerResult.Failure(OneDriveAPIError.UnspecifiedError(response)), nil, nil, nil)
                        return
                    }
                    
                    guard let expirationDateTime = jsonResponse!["expirationDateTime"] as? String else {
                        completion(OneDriveManagerResult.Failure(OneDriveAPIError.UnspecifiedError(response)), nil, nil, nil)
                        return
                    }
                    
                    guard let nextExpectedRanges = jsonResponse!["nextExpectedRanges"] as? [String] else {
                        completion(OneDriveManagerResult.Failure(OneDriveAPIError.UnspecifiedError(response)), nil, nil, nil)
                        return
                    }
                    
                    completion(OneDriveManagerResult.Success, uploadUrl, expirationDateTime, nextExpectedRanges)
                }
                catch{
                    completion(OneDriveManagerResult.Failure(OneDriveAPIError.JSONParseError), nil, nil, nil)
                }
            default:
                completion(OneDriveManagerResult.Failure(OneDriveAPIError.UnspecifiedError(response)), nil, nil, nil)
            }
        })
        task.resume()
    }
    
    
    let partSize: Int = 327680;
    func uploadBytes(uploadUrl:String, completion: @escaping (OneDriveManagerResult, _ webUrl: String?) -> Void) {
        let image = UIImage(named: "Small_Red_Rose")
        let data = UIImageJPEGRepresentation(image!, 1.0) as Data?
        let imageSize: Int = data!.count
        var returnWebUrl: String? = ""
     
        for startPointer in stride(from: 0, to: imageSize, by: partSize) {
            uploadByteParts(uploadUrl: uploadUrl, data: data!, startPointer: startPointer, endPointer: startPointer + partSize - 1, imageSize: imageSize, completion: { (result: OneDriveManagerResult, webUrl) -> Void in
                switch(result) {
                    case .Success:
                        if (webUrl?.count != 0) { returnWebUrl = webUrl }
                    case .Failure(let error):
                        completion(OneDriveManagerResult.Failure(OneDriveAPIError.GeneralError(error)), nil)
                }
            })
            //TODO: need to work away to fire these recursively waiting on the next so can return web url
            usleep(1000000)
        }
        completion(OneDriveManagerResult.Success, returnWebUrl)
    }
    
    func uploadByteParts(uploadUrl:String, data:Data,startPointer:Int, endPointer:Int, imageSize:Int, completion: @escaping (OneDriveManagerResult, _ webUrl: String?) -> Void) {
        
        var dataEndPointer = endPointer
        if (endPointer + 1 >= imageSize){
            dataEndPointer = imageSize - 1
        }
        let strContentRange = "bytes \(startPointer)-\(dataEndPointer)/\(imageSize)"
        print(strContentRange)
        
        let defaultSession = URLSession(configuration: .default)
        var request = URLRequest(url: URL(string: uploadUrl)!)
        request.httpMethod = "PUT"
        request.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")
        request.setValue("\(partSize)", forHTTPHeaderField: "Content-Length")
        request.setValue(strContentRange, forHTTPHeaderField: "Content-Range")
        
        let uploadTask = defaultSession.uploadTask(with: request, from: data[startPointer ... dataEndPointer],
           completionHandler: { (responseData, response, error) in
            if let httpResponse = response as? HTTPURLResponse {
                switch httpResponse.statusCode {
                case 200..<300:
                    print("Success")
                case 400..<500:
                    print("Request error")
                case 500..<600:
                    print("Server error")
                case let otherCode:
                    print("Other code: \(otherCode)")
                }
            }
            
            if let responseData = responseData {
                do {
                    let jsonResponse = try JSONSerialization.jsonObject(with: responseData, options: [])  as? [String: Any]
                    let webUrl = jsonResponse!["webUrl"] as? String
                    completion(OneDriveManagerResult.Failure(OneDriveAPIError.UnspecifiedError(response)), webUrl)
                }
                catch{
                    completion(OneDriveManagerResult.Failure(OneDriveAPIError.GeneralError(error)), nil)
                }
            }
            
            // Do something with the error
            if let error = error {
                completion(OneDriveManagerResult.Failure(OneDriveAPIError.GeneralError(error)), nil)
            }
        })
        uploadTask.resume()
    }

    func createFolder(folderName:String, folderId:String, completion: @escaping (OneDriveManagerResult) -> Void) {
        
        let request = NSMutableURLRequest(url: NSURL(string: "\(baseURL)/me/drive/special/approot:/\(folderName)")! as URL)
        request.httpMethod = "PUT"
        request.setValue("application/json", forHTTPHeaderField: "Content-Type")
        request.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")
        
        let emptyParams = Dictionary<String, String>()
        let params = ["name":folderName,
                      "folder":emptyParams,
                      "@name.conflictBehavior":"rename"] as [String : Any]
        
        request.httpBody = try! JSONSerialization.data(withJSONObject: params, options: JSONSerialization.WritingOptions())
        
        let task = URLSession.shared.dataTask(with: request as URLRequest, completionHandler: {
            (data, response, error) -> Void in
            
            
            if let someError = error {
                completion(OneDriveManagerResult.Failure(OneDriveAPIError.GeneralError(someError)))
                return
            }
            
            let statusCode = (response as! HTTPURLResponse).statusCode
            
            switch(statusCode) {
            case 200, 201:
                completion(OneDriveManagerResult.Success)
            default:
                completion(OneDriveManagerResult.Failure(OneDriveAPIError.UnspecifiedError(response)))
            }
        })
        task.resume()
    }

    func syncUsingViewDelta(syncToken:String?,
                            completion: @escaping (OneDriveManagerResult, _ newSyncToken: String?, _ deltaArray: [DeltaItem]?) -> Void) {
        syncUsingViewDelta(syncToken: syncToken, nextLink: nil, currentDeltaArray: [DeltaItem](), completion: completion)
    }
    
    func syncUsingViewDelta(syncToken:String?, nextLink: String?, currentDeltaArray: [DeltaItem]?,
                            completion: @escaping (OneDriveManagerResult, _ newSyncToken: String?, _ deltaArray: [DeltaItem]?) -> Void) {
        
        var currentDeltaArray = currentDeltaArray
        var request: URLRequest
        
        if let nLink = nextLink {
            request = URLRequest(url: URL(string: "\(nLink)")!)
        }
        else {
            if let sToken = syncToken {
                request = URLRequest(url: URL(string: "\(baseURL)/me/drive/root/view.delta?token=\(sToken)")!)
            }
            else {
                request = URLRequest(url: URL(string: "\(baseURL)/me/drive/root/view.delta")!)
            }
        }
        
        print("\(request)")
        
        request.httpMethod = "GET"
        request.setValue("application/json", forHTTPHeaderField: "Accept")
        request.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")
        
        let task = URLSession.shared.dataTask(with: request, completionHandler: {
            (data, response, error) -> Void in
            
            if let someError = error {
                print("error \(String(describing: error?.localizedDescription))")
                completion(OneDriveManagerResult.Failure(OneDriveAPIError.GeneralError(someError)),nil, nil)
                return
            }
            
            let statusCode = (response as! HTTPURLResponse).statusCode
            print("status code = \(statusCode)\n\n")
            
            switch(statusCode) {
            case 200:
                do{
                    let jsonResponse = try JSONSerialization.jsonObject(with: data!, options: [])  as? [String: Any]
                    print((jsonResponse?.description)!) // outputs whole JSON

                    guard let deltaToken = jsonResponse!["@delta.token"] as? String else {
                        completion(OneDriveManagerResult.Failure(OneDriveAPIError.UnspecifiedError(response)), nil, nil)
                        return
                    }
                    
                    print("delta token = \(deltaToken)")
                    
                    if let items = jsonResponse!["value"] as? [[String: AnyObject]] {
                        for item in items {
                            let fileId: String = item["id"] as! String
                            let lastModifiedRaw = item["lastModifiedDateTime"] as! String
                            let lastModified = self.localTimeStringFromGMTTime(gmtTime: lastModifiedRaw)
                            
                            let fileName: String? = item["name"] as? String
                            var isFolder: Bool
                            var isDelete: Bool
                            var parentId: String?
                            
                            if let _ = item["folder"] {
                                isFolder = true
                            }
                            else {
                                isFolder = false
                            }
                            
                            if let _ = item["deleted"] {
                                isDelete = true
                            }
                            else {
                                isDelete = false
                            }
                            
                            if let parentReference = item["parentReference"] as? [String: AnyObject] {
                                parentId = parentReference["id"] as? String!
                            }
                            
                            let deltaItem = DeltaItem(
                                fileId: fileId,
                                fileName: fileName,
                                parentId: parentId,
                                isFolder: isFolder,
                                isDelete: isDelete,
                                lastModified: lastModified)
                            
                            currentDeltaArray?.append(deltaItem)
                        }
                    }
                    
                    if let nextLink = jsonResponse!["@odata.nextLink"] as? String {
                        self.syncUsingViewDelta(syncToken: syncToken, nextLink: nextLink, currentDeltaArray: currentDeltaArray, completion: completion)
                    }
                    else {
                        completion(OneDriveManagerResult.Success, deltaToken, currentDeltaArray)
                    }
                }
                catch{
                    completion(OneDriveManagerResult.Failure(OneDriveAPIError.JSONParseError), nil, nil)
                }
                
            default:
                completion(OneDriveManagerResult.Failure(OneDriveAPIError.UnspecifiedError(response)), nil, nil)
            }
        })
        
        task.resume()
    }
    
    func localTimeStringFromGMTTime(gmtTime: String) -> String {
        
        let locale = Locale(identifier: "en_US_POSIX")
        let dateFormatterFrom = DateFormatter()
        dateFormatterFrom.locale = locale as Locale!
        dateFormatterFrom.timeZone = TimeZone(abbreviation: "GMT")
        dateFormatterFrom.dateFormat = "yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'"
        
        let lastModifiedDate = dateFormatterFrom.date(from: gmtTime)
        
        let dateFormatterTo = DateFormatter()
        dateFormatterTo.locale = locale as Locale!
        dateFormatterTo.timeZone = TimeZone.current
        dateFormatterTo.dateFormat = "yyyy'-'MM'-'dd' 'HH':'mm':'ss"
        
        return dateFormatterTo.string(from: lastModifiedDate!)
    }
    
}









