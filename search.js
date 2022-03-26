function wildcardMatches(str, rule) {
    var currentLoc = 0;
    for (var i = 0; i < rule.length; i++) {
        var loc = str.indexOf(rule[i], currentLoc);
        if (loc < 0)
            return false;
        if ((!i) && (rule[0] != ""))
            if (loc)
                return false;
        if ((i == rule.length - 1) && (rule[i] != ""))
            if (loc + rule[i].length != str.length)
                return false;
        currentLoc = loc + rule[i].length;
    }
    return true;
}

function lookInFolder(fld)
{
    for (var e = new Enumerator(fld.Files);!e.atEnd();e.moveNext())
    {
        var file = e.item();
        if (wildcardMatches(file.Name.toLowerCase(), r))
            window.handleResult(file.Path);
    }

    for (var e = new Enumerator(fld.SubFolders);!e.atEnd();e.moveNext())
        lookInFolder(e.item());
}

var fso = new ActiveXObject("Scripting.FileSystemObject");
if (fso.FolderExists(tparam.searchIn))
{
    var r = tparam.fileName.toLowerCase().split("*");
    lookInFolder(fso.GetFolder(tparam.searchIn));
}
else
    window.alert("Unable to find the folder '" + tparam.searchIn + "'.");

window.searchCompleted();