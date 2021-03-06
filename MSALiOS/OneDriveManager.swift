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
    
    func createSharingLink(fileId:String,
                        completion: @escaping (OneDriveManagerResult, _ webUrl: String?) -> Void) {
        
        let request = NSMutableURLRequest(url: URL(string: "\(baseURL)/me/drive/items/\(fileId)/createLink")!)
        
        request.httpMethod = "POST"
        request.setValue("application/json", forHTTPHeaderField: "Content-Type")
        request.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")
        request.httpBody =  ("{\"type\": \"view\",\"scope\": \"anonymous\"}" as NSString).data(using: String.Encoding.utf8.rawValue)
        
        let task = URLSession.shared.dataTask(with: request as URLRequest, completionHandler: {
            (data, response, error) -> Void in
            
            guard error == nil else {
                print("error calling upload")
                print(error!)
                return
            }
            guard let responseData = data else {
                print("Error: did not receive data")
                return
            }
    
            do{
                let json = try JSONSerialization.jsonObject(with: responseData, options: JSONSerialization.ReadingOptions()) as! [String:Any]
                print((json.description)) // outputs whole JSON
                
                let decoder = JSONDecoder()
                let sharingLinkRespObj = try decoder.decode(SharingLinkRespObj.self, from: responseData)
                
                let webUrl = sharingLinkRespObj.link.webUrl
                completion(OneDriveManagerResult.Success, webUrl)
            }
            catch{
                completion(OneDriveManagerResult.Failure(OneDriveAPIError.JSONParseError), nil)
            }
        })
        
        task.resume()
    }
    
    
    /*
     [
        "link": {
                     application =     {
                     id = 4c211b04;
                 };
         type = view;
         webUrl = "https://1drv.ms/u/s!AuIrNnKy4gormcI5WGgWHKcB5-2YZA";
         }, "@odata.type": #microsoft.graph.permission, "id": ZKyb1FYbYhklX74gxTQ3IaFRuXE, "roles": <__NSSingleObjectArrayI 0x60000001a520>(
                         read
                         )
     , "shareId": s!AuIrNnKy4gormcI5WGgWHKcB5-2YZA, "@odata.context": https://graph.microsoft.com/v1.0/$metadata#permission]
     */
    struct SharingLinkRespObj : Decodable {
        let id: String
        let roles: [String]?
        let link: LinkRespObj
    }
    struct LinkRespObj : Decodable {
        let webUrl:String?
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
    func uploadBytes(uploadUrl:String, completion: @escaping (OneDriveManagerResult, _ webUrl: String?, _ fileId: String?) -> Void) {
        let image = UIImage(named: "Small_Red_Rose")
        let data = UIImageJPEGRepresentation(image!, 1.0) as Data?
        let imageSize: Int = data!.count
        var returnWebUrl: String?
        var returnFileId: String?
        var returnNextExpectedRange: Int = 0
        
        let dispatchGroup = DispatchGroup()
        let dispatchQueue = DispatchQueue(label: "taskQueue")
        let dispatchSemaphore = DispatchSemaphore(value: 0)
        
        let urlSessionConfiguration: URLSessionConfiguration = URLSessionConfiguration.default.copy() as! URLSessionConfiguration
        urlSessionConfiguration.httpMaximumConnectionsPerHost = 1
        let defaultSession = URLSession(configuration: urlSessionConfiguration)
        
         let uploadBytePartsCompletionHandler: (OneDriveManagerResult, Int?, String?, String?) -> Void = {
             (result: OneDriveManagerResult, nextExpectedRange, webUrl, fileId) in
             switch(result) {
                 case .Success:
                    if (nextExpectedRange != nil) {returnNextExpectedRange = nextExpectedRange!}
                    if (webUrl != nil) { returnWebUrl = webUrl}
                    if (fileId != nil) { returnFileId = fileId}
                    dispatchSemaphore.signal()
                    dispatchGroup.leave()
                 case .Failure(let error):
                    completion(OneDriveManagerResult.Failure(OneDriveAPIError.GeneralError(error)), nil, nil)
             }
         }
        
        // we need a dispatch queue to handle the async wait pattern on calling the graph.microsoft.com api
        dispatchQueue.async {
            // this will set it up to call recursively, as the graph.microsoft.com API response gives you the next start pointer to send which we update in the completion handler
            while (returnWebUrl == nil) {
                dispatchGroup.enter()
                self.uploadByteParts(defaultSession: defaultSession, uploadUrl: uploadUrl, data: data!, startPointer: returnNextExpectedRange, endPointer: returnNextExpectedRange + self.partSize - 1, imageSize: imageSize, completion: uploadBytePartsCompletionHandler)
                dispatchSemaphore.wait()
            }
        }
        dispatchGroup.notify(queue: dispatchQueue) {
            DispatchQueue.main.async {
                completion(OneDriveManagerResult.Success, returnWebUrl, returnFileId)
            }
        }
    }
    
    func uploadByteParts(defaultSession: URLSession, uploadUrl:String, data:Data,startPointer:Int, endPointer:Int, imageSize:Int, completion: @escaping (OneDriveManagerResult, _ nextExpectedRangeStart: Int?, _ webUrl: String?, _ fileId: String?) -> Void) {
        
        var dataEndPointer = endPointer
        if (endPointer + 1 >= imageSize){
            dataEndPointer = imageSize - 1
        }
        let strContentRange = "bytes \(startPointer)-\(dataEndPointer)/\(imageSize)"
        print(strContentRange)
        
        var request = URLRequest(url: URL(string: uploadUrl)!)
        request.httpMethod = "PUT"
        request.setValue("Bearer \(accessToken)", forHTTPHeaderField: "Authorization")
        request.setValue("\(partSize)", forHTTPHeaderField: "Content-Length")
        request.setValue(strContentRange, forHTTPHeaderField: "Content-Range")
        
        let uploadTaskCompletionHandler: (Data?, URLResponse?, Error?) -> Void = {
            (data, response, error) in
            
            guard error == nil else {
                print("error calling upload")
                print(error!)
                return
            }
            guard let responseData = data else {
                print("Error: did not receive data")
                return
            }
            do {
                guard let json = try JSONSerialization.jsonObject(with: responseData, options: []) as? [String: Any] else {
                    print("error trying to convert data to JSON")
                    return
                }
                print("The json is: " + json.description)
                
                guard let webUrl = json["webUrl"] as? String else {
                    let decoder = JSONDecoder()
                    let uploadTaskObj = try decoder.decode(UploadTaskObj.self, from: responseData)
                    
                    let strNextExpectedRanges = uploadTaskObj.nextExpectedRanges![0]
                    let index = strNextExpectedRanges.index(of: "-")!
                    let strNextExpectedRangeStart = strNextExpectedRanges.substring(to: index)
                    completion(OneDriveManagerResult.Success, Int(strNextExpectedRangeStart), nil, nil)
                    return
                }
                
                let fileId = json["id"] as? String
                completion(OneDriveManagerResult.Success, nil, webUrl, fileId)
            } catch  {
                print("error trying to convert data to JSON")
                completion(OneDriveManagerResult.Failure(OneDriveAPIError.GeneralError(error)), nil, nil, nil)
            }
        }
        
        let uploadTask = defaultSession.uploadTask(with: request, from: data[startPointer ... dataEndPointer], completionHandler: uploadTaskCompletionHandler)
        uploadTask.resume()
    }
    
    /*
     ["expirationDateTime": 2018-04-19T16:23:12.576Z, "nextExpectedRanges": <__NSSingleObjectArrayI 0x60c00001c3d0>(
     327680-2718924
     )
     ]
     */
    struct UploadTaskObj : Decodable {
        let expirationDateTime: String
        let nextExpectedRanges: [String]?
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









