// Taking a look at this code still reminds me of Omegastripes's post on Stack Overflow,
// which illustrated how to use WSH VBS to dynamically create an HTA window, and receive
// its window object through a Web Browser Control. Here is its link:
// https://stackoverflow.com/questions/47100085/creating-multi-select-list-box-in-vbscript/47111556#47111556
// Thanks a lot to his post, I'm using a similar mechanism in this component to
// transfer thread data (like the thread template element, the extra parameter,
// the thread flags, .etc) from the mshta.exe process (The HTA process) to the
// wscript.exe process (the underlying process of the thread) via a dedicated
// Web Browser Control (WBC). Ha ha!

eval((function() {
    var TF_SUSPENDED        = 1;
    
    var e = new Enumerator(WScript.Arguments);
    if (e.atEnd())
        throw new Error("Not enough arguments");
    
    var signature = "app" + e.item();
    var wins = WScript.CreateObject("Shell.Application").Windows();
    var threadsInfo;
    
    for (var i = 0;i < wins.Count;i++)
    {
        var win = wins.Item(i);
        if (win)
        {
            threadsInfo = win.GetProperty(signature);
            if (threadsInfo)
            {
                win.RegisterAsBrowser = false;
                break;
            }
        }
    }
    
    if (!threadsInfo)
        throw new Error("The property '" + signature + "' does not exist in any windows.");
    
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
