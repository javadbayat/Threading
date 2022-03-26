# Threading
An HTML component which enables you to create threads in HTAs, along with an example of multi-threaded app that makes use of this component.

# Example app
Run the app by double-clicking the file 'multithreading example.hta'. The app has two tabs: **Search** and **Mix Images**.

## Search tab
In the Search tab, you can search your file system. In the **Location** field, enter the path of the directory in which the app will look for files. In the **File name** field, enter the name of the file you wish to find. Note that you can also place *wildcards* in the file name.

Finally, click the **Search** button to start the search. Technically, this will result in creation of a thread that performs the search operation. After the search is started, you can stop it by simply clicking the **Stop** button. Technically, this will result in the search thread being killed.

The results of the search will be displayed in the light-yellow box below the buttons. Clicking on each result will make the app open the containing folder of the file in a File Explorer window. You can also clear the list of results by clicking the **Clear Results** button next to the Search button.

Clicking **Read Results** button will make the app create a thread to read the search results aloud via the system's text-to-speech engine. After clicking it, you can stop the playback by clicking it again.

## Mix Images tab
In the Mix Images tab, the app takes two images as input and mixes them to produce a couple of composite images each of which contains the input images overlaid with each other. So it's a *magical operation*! Try it to see what it really means.

In the **Input** section, click the **Browse** buttons to select the first and second input image files that are to be mixed.  
In the **Output** section, click the **Browse** button to select the directory in which the two composite images will be saved. The name of these two composite image files will be in the following format by default:  
*n*.bmp  
where n is an automatically generated number. For example, firstly, the name '1.bmp' is chosen by the app. But if the file '1.bmp' already exists, the name '2.bmp' is chosen; and if it again exists, '3.bmp' is chosen, and so on.

Now start mixing images by clicking the **Mix!** button. Watch the progress bar until the operation is completed, and the following message pops up:

> The images were successfully mixed!

After the mix operation, these two composite image files will be approximately identical to each other. But wait - do not assume that the second file is useless! You are able to *reconstruct* the original input image files from these composite image files if both of them are present. To do so, you need to start a new mix operation with the app by specifying the same composite images as input images. Then the app will create output images that are identical to the original images!

Then what is the **Unit size** field in the Settings section? It's hard to explain, but we recommend that you leave it be the default (1). Because the higher its value, the lower the quality of the composite images.

# thread.htc Component
**thread.htc** is an HTML component (HTC file) that enables the creation of ***virtual threads*** in HTML Applications (HTAs). The term "virtual" means that these threads are not really threads, but are actually *wscript.exe* processes that communicate with the HTA process (mshta.exe) via COM. Moreover, the code that is executed by these threads is originally stored within the HTA file, and as soon as the thread starts, the code is dynamically transfered to the 'wscript.exe' process for execution.

To use the component in an HTA, the first step is to declare an XML namespace named 't', and then import 'thread.htc' into that namespace [(More about this here)](https://docs.microsoft.com/en-us/previous-versions//ms531426(v=vs.85)?redirectedfrom=MSDN).

Then you must define one or more thread templates, which contain the JScript code to be executed by a thread. Then you can create from the thread template as many threads as desired, all of which execute the code within the template. Use the `<t:thread>` element to define a thread template. This element must necessarily have an id. For example, the following piece of code defines a template for a thread that simply moves a large file from one location to another:

    <t:thread id="fileMoverThread">
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    fso.MoveFile("C:\\movie.mp4", "D:\\movie.mp4");
    </t:thread>

## To be continued...
