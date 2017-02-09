# MaM.Hoard.Organizer
A PS tool to organize way, way, way too many ebooks.
I have a different script for Audiobooks, but will merge both of them into this one later down the road.

Essentially, this script works by searching MaM for an eBook and then moving the file to its appropriate folder that MaM categorizes it as. This is helpful if you have a ton of ebooks downloaded from other places, and want to use MaM's excellent categorizing scheme to organize it. 

========================================================================================================================================

How to set it up:
1) Create a bunch of folders in a directory that correspond to the categories in MaM. I my case, I used a bang(!) + category name, which makes it easy to tell which folders are organized and which aren't. If you want to match my directory structure to get started, take a look at the $dirtable hashable in the script. Laer I'll add functionality to create the directories if you want them.
2) Set the $Source and $Dest parameters in the script. $Source is where the uncategorized files are, $Dest is the folder with your new categories.
3) The first time you run the script, you will need to log in to MaM with IE. This can be done by just running these two commands in a PS window prior to running the script:

$ie = new-object -ComObject "InternetExplorer.Application"
$ie.visible = $true

That's it. Run that bad boy.

Running this in the future, you can change the "$ie.visible" to $false and an IE window won't appear when you are running the script. 

========================================================================================================================================

Now, there are a few limitations of the script. It isn't perfect.
One thing to note is that sometimes there wil be a false alarm for categorizing books under "Mixed Collections." For example, I've found that the script will cycle through, for example, 10 books in a series, and categorize the first 5 as Science Fiction but categorize the next few as Mixed Collections, which can be annoying. To help with this, whennever the script searches it will always only return the two top results, and if both of them are Mixed Collections it throws the category out.

That's the only real big thing here. Otherwise it seems to work pretty well. One other thing to note is this script does not do comics, as the MaM "Comics" category is not very specific.  
