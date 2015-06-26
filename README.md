# EmojiO #
EmojiO is an Outlook Add-In that brings the fun of using Emoji to Outlook.  Deeply integrated into the Office User Interface EmojiO is at the users fingertips right on the Outlook Ribbon. 

![](http://i.imgur.com/ZostAmO.png)

## User Features ##

- **Quick Insert**. Click, click and you have an emoji inserted.
- **Gallery**. Dozens of emoticons to chose from with real-time search built-in
- **Emojify selection**. Parses the selected text and replaces any emoji codes or text smiles with their corresponding images. 

## Installation ##
If all you are interested in is to see the Add-In in action you can install it in a snap. Here are the steps:

1. Ensure you have the latest [Office 2016 Preview](https://products.office.com/en-us/office-2016-preview). The add-in also works with Office 2013 but the feature to show it as buttons on the Ribbon is only currently available in the Office 2016 preview. 
2. Register the Add-in into your mailbox. This simply means registering the manifest file with your mailbox. This [video](https://www.youtube.com/watch?v=jQGDZNsCkQs) shows you how to register an add-in in your personal mailbox. The mailbox must be an **Exchange mailbox** because that is where mail Add-Ins are registered.  
3. Done!

## Under the hood ##
Emojio is built as an Office Web Add-In and leverages the following platform features:

- [Add-In Commands](http://aka.ms/addincommands) to project user interface elements in Office
- The [MailboxAPI](https://msdn.microsoft.com/EN-US/library/office/dn705877.aspx) of Office.js to interact with mail items
- Office UI Fabric to provide some of the visuals (e.g. the SearchBox)
- [Azure Web App Servic](http://azure.microsoft.com/en-us/services/app-service/web/)e to host the underlying web app
- [EmojiOne](http://emojione.com/) for all the art. Thank you EmojiOne!

## Source Code ##
All the key pieces you need to are part of this project. The key files are the [manifest](https://github.com/LezaMax/emoji/blob/master/manifest/emojiManifest.xml) and the [web page](https://github.com/LezaMax/emoji/blob/master/webapp/emojiPane.html) that powers the experience. 

