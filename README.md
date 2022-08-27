# folderSyncDemo
Windows service project to sync operation and permission management. 

This project has been done to sync between two Windows directories. As new files arrived into one directory and that directory is not acccessible for all the users due to security issues, I had to built this tool to sync that directory with another accessible one. 

So basically, this is a windows service that runs on background and checkany changes on the targetted directory. If chnaged then just do the same stuff with the directory that needs to be updated. It only copies or delete files that is changed not just delete all and copy. 
