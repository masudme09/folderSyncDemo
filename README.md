# FolderSync
**Windows service project to sync operation and permission management.** 

The purpose of this project is to implement a Windows service for synchronizing the operations and permission management between two directories. Due to security concerns, one of the directories is inaccessible to certain users, which necessitated the development of this tool to facilitate synchronization between the two.
Essentially, the service runs in the background and continuously monitors the targeted directory for any modifications. When a change is detected, the tool applies the same operation to the directory that requires updating. Rather than performing a complete deletion and copy of all files, the tool selectively copies or deletes only the files that have been modified. 
