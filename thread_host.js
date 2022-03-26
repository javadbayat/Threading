eval((function() {
    var TF_SUSPENDED        = 1;
    
    var e = new Enumerator(WScript.Arguments);
    if (e.atEnd())
        throw new Error("Not enough arguments");
    
    var appid = "app" + e.item();
    var wins = WScript.CreateObject("Shell.Application").Windows();
    var threadsInfo;
    
    for (var i = 0;i < wins.Count;i++)
    {
        var win = wins.Item(i);
        if (win)
        {
            threadsInfo = win.GetProperty(appid);
            if (threadsInfo)
            {
                win.RegisterAsBrowser = false;
                break;
            }
        }
    }
    
    if (!threadsInfo)
        throw new Error("The property '" + appid + "' does not exist in any windows.");
    
    e.moveNext();
    if (e.atEnd())
        throw new Error("Not enough arguments");
    
    var tmid = Number(e.item());
    thread = threadsInfo(tmid);
    tparam = thread.arg;
    window = thread.element.document.parentWindow;
    
    thread.script = this;
    thread.suspend = function() {
        thread.flags |= TF_SUSPENDED;
        do
            WScript.Sleep(1000);
        while (thread.flags & TF_SUSPENDED)
    };
    
    if (thread.flags & TF_SUSPENDED)
        thread.suspend();
    
    return thread.element.innerHTML;
})());