# Image to link converter for excel

This script will find all the images in a file and convert them to their hyperlinks (if attached)

The script has two settings that can be toggled by variable: 
- replace image 
- add column next to image.

## Motivation
Sometimes you copy a list from the web that contains images with links.  
This is a problem when you want to process the files using a library like pandas. 
These libraries indentify an object but not the link in a deep copy.
This means you loose critical informaiton during processing.
I couldn't find any other code online that does this and it isn't supported by pandas or xls reader 
This code means that you preserve the links on images.
