from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive

# from pydrive2 import drive as somedrive
from send2trash import send2trash
from threading import Timer
from pprint import pprint
import os
import json

from odf_manip import magicParse, customClassify, replaceLinks
"""
metadata_format = {
        "type": ("source", "note", "position", "resolution", "unclassified"),
        "agenda": "string",
        "committee": "string",
        "country": "string",
        "filetype": ("[m]html", "pdf", "docx", "googledoc", "md", "org", "etc."),
}
"""


def authorisedDrive():
    """

    :returns: An authorised google drive object which can be used to interact with user's drive
    :rtype: pydrive2.drive.GoogleDrive object

    """
    gauth = GoogleAuth()
    return GoogleDrive(gauth.LocalWebserverAuth(
    ))  # Creates local webserver and auto handles authentication.


def deAuthorise():
    """Trash the credential files, so the program needs to be re-authorised.
    :returns: None
    :rtype: NoneType

    """
    send2trash("credentials.json")  # NOTE: Lazy as hell.


mydrive = authorisedDrive()


def getFile(filename, drive=mydrive):
    """Search the user's drive for a file with a given name and return it

    :param filename: The name to search for in the google drive
    :param drive: The GoogleDrive where we should search
    :returns: A single google drive file, which matches the filename
    :rtype: driveFile object

    """
    queryParams = {
        # "corpora": "user",
        "q": f"title contains '{filename}' and trashed = False"
        # "maxresults": 3,
    }
    files = drive.ListFile(queryParams).GetList()
    return files[0]


def getAllFiles(filename, drive=mydrive):
    """Search the user's drive for a file with a given name and return all files found

    :param filename: The name to search for in the google drive
    :param drive: The GoogleDrive where we should search
    :returns: A list of google drive files, which match the filename
    :rtype: list of driveFile objects

    """
    queryParams = {
        # "corpora": "user",
        "q": f"title contains '{filename}' and trashed = False"
        # "maxresults": 3,
    }
    files = drive.ListFile(queryParams).GetList()
    return files


def getExistingFolder(filename, parentId, drive=mydrive):
    """Returns a google drive folder which matches the filename if one is found, otherwise false. Used as a check

    :param filename: The folder name to search for
    :param parentId: The ID of the folder in which you are searching
    :param drive:
    :returns: Google drive folder with that filename, or False
    :rtype: driveFile, or boolean

    """
    queryParams = {
        # "corpora": "user",
        "q":
        f"title='{filename}' and mimeType='application/vnd.google-apps.folder' and '{parentId}' in parents and trashed = False"
        # "maxresults": 3,
    }
    files = drive.ListFile(queryParams).GetList()
    return files[0] if files else False


def downloadHelper(fileObj, appname="pyMUN"):
    """A wrapper function that downloads a google drive file locally, saves it to the right path, and returns some useful data about the file

    :param fileObj: The DriveFile object to download
    :param appname: The name of the subfolder in which to store the document
    :returns: A dictionary with the path, id, name, and mimetype of the document
    :rtype: Dict

    """
    """
        We define a helper/wrapper that downloads a file to the right place, and returns a dict/tuple of:
        - File path
        - ID
        - Title
        - (Metadata from the description): All of it
        - (Some things from the fileObj dict): mimeType as originalMime
        """

    # We assume that we actually have to download this, so the metadata check is already done
    path = f"{os.path.expanduser('~')}/tmp/{appname}/{fileObj['id']}.docx"
    fileObj.GetContentFile(path)
    return {
        "path": path,
        "id": fileObj["id"],
        "name": fileObj["title"],
        "originalMime": fileObj["mimeType"],
    }


# We're using the description to store metadata as JSON, since I can't make the properties work.
# Implies we have to insert some basic checks for corrupt data: If we face an err reading JSON we should just replace it and start over.


def makeDriveFile(localpath, drive=mydrive):
    """Uploads a modified docx file back to google drive

    :param localpath: The path to the document
    :param drive: The drive object used for uploading
    :returns: A drive file object representing a word document
    :rtype: DriveFile object

    """
    f = drive.CreateFile({"id": localpath.split("/")[-1].replace(".docx", "")})
    f.SetContentFile(localpath)
    return f


def getMetadata(fileObj):
    """Return the metadata we've set for the fileobject in it's 'description' section

    :param fileObj: The drive file to query
    :returns: Dict representing metadata, which may or may not be empty
    :rtype: dict

    """
    try:
        return json.loads(fileObj["description"])
    except:
        return dict()


def setMetadata(fileObj, dataDict):
    """Sets the description attr of the google drive file to a json dump of the dataDict. Used as our personal metadata store, overwriting existing metadata.

    :param fileObj: The drive file who's metadata we want to set
    :param dataDict: The metadata we want to declare, in the form of a dictionary
    :returns: A file object with that metadata(description) added on
    :rtype: DriveFile object

    """
    fileObj["description"] = json.dumps(dataDict)
    return fileObj


# We might have to perform several functions on a doc before uploading it. So it makes more sense to return than to upload, for performance


def addMetadata(fileObj, dataDict):
    """Used to add specific metadata to a drive file without overwriting existing

    :param fileObj: The drive file in question
    :param dataDict: The data we want to add
    :returns: drive file with updated metadata
    :rtype: DriveFile object

    """
    # Uses calls to get and set. NOTE: set/get are our expensive ops. So try to minimise them
    # Should override existing keys, if called.
    meta = getMetadata(fileObj)
    meta.update(dataDict)
    return setMetadata(fileObj, dataDict)


def deleteMetadata(fileObj, keyArray=[]):
    """Deletes some or all existing metadata of a drive file

    :param fileObj: The DriveFile object which is the target.
    :param keyArray: If present, function only deletes those specific keys.
    :returns: An updated drive file
    :rtype: DriveFile object

    """
    # Default to deleting all metadata
    if keyArray:
        meta = getMetadata(fileObj)
        for i in keyArray:
            meta.pop(i, None)
        return setMetadata(fileObj, meta)
    else:
        return setMetadata(fileObj, dict())


def createLink(fileObj, folderObj, drive=mydrive):
    """Creates a link to/copy of the given drive file in the given drive folder

    :param fileObj: The file to which we want to create a link
    :param folderObj: The folder in which the link should be placed
    :returns: The file object with a modified 'parents' attribute
    :rtype: DriveFile object

    """
    fileObj["parents"].append({
        "kind": "drive#parentReference",
        "id": folderObj["id"],
        "selfLink":
        f"https://www.googleapis.com/drive/v2/files/{fileObj['id']}/parents/{folderObj['id']}",
        "parentLink":
        f"https://www.googleapis.com/drive/v2/files/{folderObj['id']}",
        "isRoot": False,
    })
    return fileObj


# Not sure what is_root:true does, so I removed it for now. Put it back if something breaks


def createFolder(name, parentId, drive=mydrive):
    """Creates and returns a drive folder within the path specified by parentId, and with the given name. If one exists already, return that instead

    :param name: The name of the folder
    :param parentId: The Id of the folder in which the new folder should be created
    :param drive: The GoogleDrive object representing a drive to modify
    :returns: An existing folder, if one is found, else a new one with the specified name and parent
    :rtype: DriveFile object

    """
    # If a folder with that name already exists at that path, just return that instead of creating a new on# If a folder with that name already exists at that path, just return that instead of creating a new one
    folderMeta = {
        "title": name,
        # The mimetype defines this new file as a folder, so don't change this.
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [{
            "kind": "drive#parentReference",
            "id": parentId
        }],
    }
    x = getExistingFolder(name, parentId, drive)
    return x if x else drive.CreateFile(folderMeta)


def getMimeType(fileObj):
    """

    :param fileObj: DriveFile object
    :returns: Mimetype of the given file
    :rtype: String

    """
    return fileObj["mimeType"]


def mimeToName(mime):
    """Given a mimetype, return the corresponding colloquial name based on a conversion table

    :param mime: Mimetype, as a string
    :returns: A colloquial name for that kind of file
    :rtype: String

    """
    # Goes from formal mimetype application/pdf to simple pdf
    conversionTable = {
        "application/vnd.google-apps.document": "gdoc",
        "application/pdf": "pdf",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        "word",
        "text/html": "html",
        "message/rfc822": "mhtml",
        "application/x-mimearchive": "mhtml",
    }
    return (conversionTable[mime]
            if mime in conversionTable.keys() else mime.split("/")[-1])


def classifyFile(fileObj):
    """Classify a non-word file.

    :param fileObj: Drive File object, representing a non-word file
    :returns: One of ("source","note","unclassified")
    :rtype: String

    """
    mime = mimeToName(getMimeType(fileObj))
    if "html" in mime or "pdf" in mime:
        return "source"
    if "markdown" in mime or "plain" in mime:
        return "note"
    return "unclassified"


def updateMetadata(fileObj):
    """Download and parse the file to identify the type, etc. and then update the fileObj with the requisite metadata, and with the links replaced

    :param fileObj: The file object to download and analyse
    :returns: A drive file with the requisite metadata added onto it
    :rtype: DriveFile object

    """
    # Updates the metadata based on reading the file and stuff
    mime = getMimeType(fileObj)
    filetype = mimeToName(mime)
    result = {"filetype": filetype}  # DONE: Fill out in requisite format
    # Download google doc via downloadHelper
    localMeta = downloadHelper(fileObj)
    replaceLinks(localMeta["path"])
    fileObj.SetContentFile(localMeta["path"])
    if filetype in ("gdoc", "word"):
        # Do the docx parsing magic on that doc, convert the return values into metadata.
        print(magicParse(localMeta["path"]))
        result.update(magicParse(localMeta["path"])
                      )  # DONE Should get type, agenda, committee, country
        custom = customClassify(localMeta["name"], localMeta["path"])
        if custom:
            result.update({"type": custom})
        # Overwrite if a custom rule takes precedence
    else:
        result.update({"type": classifyFile(fileObj)})
    send2trash(localMeta["path"])
    return addMetadata(fileObj, result)


# The linking/copying files so they show up in multiple folders is based on messing with the parent attr of the file object
# https://developers.google.com/drive/api/v2/reference/files

# Resolution data structure:
# clause = ["Urges .... inter alia:", ["<Another clause>"]]
# So basically, it's a tree. Each clause is a root with a certain amount of children.

# NOTE: getMainFolder and createTypeFolders are expensive and shouldn't be repeated unduly, so makes sense to save as global vars


def getChild(name, parentId, drive=mydrive):
    """Search within a folder for a given file, and return it if found

    :param name: The name of the folder to search for
    :param parentId: The Id of the folder in which to search
    :param drive: The GoogleDrive object
    :returns: The child folder which has that name
    :rtype:DriveFile object

    """
    queryParams = {
        "q":
        f"title = '{name}' and '{parentId}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed=False",
    }
    files = drive.ListFile(queryParams).GetList()
    return files[0]


def getMainFolder(path, drive=mydrive):
    """Based on a path (Unix style, with /), find the folder in the user's google drive represented by that

    :param path: The path which represents a particular drive folder. Use Unix syntax, not MS
    :param drive: GoogleDrive object in which to search
    :returns: The google drive folder at the specified path
    :rtype: DriveFile object

    """
    pathElems = (path.split("/") if path.split("/")[0] else path.split("/")[1:]
                 )  # Ignore the first element, which should be empty
    child = {"id": "root", "alternateLink": "https://drive.google.com"}
    if not any(pathElems):  # If it doesn't have any non-empty strings
        return child
    for i, v in enumerate(pathElems):  # Index, val
        child = getChild(pathElems[i], child["id"], drive)
    return child


def createTypeFolders(root,
                      types=("source", "note", "position", "resolution",
                             "unclassified")):
    """Creates the folders to store/sort different kinds of documents

    :param root: The ID of the folder in which these new folders should be created
    :param types: A tuple/list of document types, for which folders should be created
    :returns: A dictionary where keys are the types, and the vals are the folders which represent those types
    :rtype: dict (keys=strings, vals=DriveFile objects)

    """
    folders = [createFolder(i, root) for i in types]
    for i in folders:
        i.Upload()

    return dict(zip(types, folders))


def listFiles(root, drive=mydrive):
    """Like ls for google drive, lists all the children of a given folder.

    :param root: The folder object in question
    :param drive: GoogleDrive object
    :returns: A list of files, all of which are children of the specified root
    :rtype: List (elems=DriveFile objects)

    """
    queryParams = {
        "q":
        f"'{root['id']}' in parents and mimeType != 'application/vnd.google-apps.folder' and trashed=False",
    }
    files = drive.ListFile(queryParams).GetList()
    return files


def updateAllMetadata(files):
    """Analyse all the files given, and update their metadata accordingly

    :param files: The list of files to download and analyse
    :returns: A list of file objects with updated metadata
    :rtype: List (elems=DriveFile objects)

    """
    # File list is optional param, so we can update selective files if we have to. For instance, only add the metadata we have to, and only sort those rather than the whole list
    toUpdate = [i for i in files if not getMetadata(i)]
    return [updateMetadata(i) for i in toUpdate]


def sortIntoFolder(fileObj, types):
    """

    :param fileObj: The file object, with metadata on it, which needs to be sorted
    :param types: A dict, as returned by createTypeFolders
    :returns: A file object representing a copy/link of the original file, in the correct folder
    :rtype: DriveFile object

    """
    # Non-destructive. It adds to the list of parents, but does not replace anything.
    meta = getMetadata(fileObj)
    doctype = meta["type"]
    return createLink(fileObj, types[doctype])


def sortAllFiles(files, types):
    """Sort all files into folders based on their updated metadata

    :param files: A list of files, who's metadata has been updated
    :returns: A list of updated files, sorted into the requisite folders
    :rtype: List (elems=DriveFile objects)

    """
    # NOTE: Works for now but performance will be hell. So ideally try to find a quicker way to do this, checking which files are uploaded already. In the long run, I can try implementing a cache and the like.
    return [sortIntoFolder(i, types) for i in files]


# So far we only have 1 upload call in a function. This is good, since everything else simply returns. We can pretty easily slap on an upload() method call when we actually call the functions:
"""TODO:
- Break an odt/odf document into a nice tree
- Mess with the tree
- Write the tree back to a document
- Consider adding an option for users to call for a hard refresh
- Figure out why stuff isn't being saved to folders
- Auto-truncate data, to prevent it from picking up too much crud.
Saving documents: Save it to ~/tmp/mundoc/<id>.odt. We'll use the ID as the unique identifier as much as possible. TODO: Handle reuploading modified files, if and when I implement auto-formatting
"""


def batchProcess(drive=mydrive):
    """A single function that processes all files in the drive, sorts them, etc. as appropriate. Should be automatically run regularly

    :param drive: The drive object in which to look
    :returns: None
    :rtype: NoneType

    """
    mainFolder = json.load(open("config.json"))["folderpath"]
    mainFolder = getMainFolder(mainFolder)
    types = createTypeFolders(mainFolder["id"])
    relevant = listFiles(mainFolder)
    updated = updateAllMetadata(relevant)
    sortedFiles = sortAllFiles(updated, types)
    for i in sortedFiles:
        i.Upload()


def main():
    # I think 'run every 30-60 minutes is a good median
    seconds = json.load(open("config.json"))["delay"] * 60
    batchProcess()
    Timer(seconds, main).start()


# if __name__ == "__main__":
#    main()
batchProcess()
# x = getAllFiles("Hello.txt")
# print(len(x))
# pprint(x[0])
# print(x[0]["title"])
# y = getExistingFolder("unclassified", "root")
# pprint(listFiles(y))
# z=createLink(x, y)
# z.Upload()
# batchProcess()                  #
"""
DONE mainFolder is a slight issue, all the rest are manageable
DONE sortAllFiles and its dependencies are issues since they require a well-defined type dict
DONE We have a preliminary doctype defined, and now the task is parsing the document and extracting good metadata

Auto-formatting will be ignored for the moment.
File metadata is unlikely to change over time. So when we run 'batchprocess', we download files with either
a) no metadata
b) missing metadata (i.e if the doctype is not clarified. The other stuff is tricky to get)
"""
