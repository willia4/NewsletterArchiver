 A simple C# "script" to scan an inbox via IMAP for unread messages from a certain sender and then compile those unread messages into an epub. 
 
#### Executing
Execute this "script" by running `dotnet run` in the same directory that contains this readme file.

#### Configuration

There are six configuration options: 
* `MailServer`: The IMAP mail server to query
* `MailUser`: The username to use to authenticate against the IMAP mail server
* `MailPassword`: The password to use to authenticate against the IMAP mail server

* `SearchFrom`: The string to use when querying the IMAP server's "From" field. Also serve's as the epub's author.
* `BookTitle`: The title of the output ePub
* `SubjectRegex`: [optional] a .NET regular expression which, if present, will be applied to each mail's subject and replaced with an empty string before using that subject to create a chapter title 


If you want to be able to just run `dotnet run`, you can add a file named `appsettings.secrest.json` at the root of this source directory with these settings. This file is excluded from git by the `.gitignore`.

```json
{
  "MailServer": "mail.example.com",
  "MailUser": "groot",
  "MailPassword": "password",
  
  "SearchFrom": "James Kirk",
  "BookTitle": "A Memoir of the Stars",
  "SubjectRegex": "A Memoir of the Starts:\\s+?"
}
```

If you don't want to write your email password to disk, you can pass some or all of these in at the command line. 

```shell
dotnet run -- --MaiLServer "mail.example.com" --MailUser "groot" --MailPassword "password" --SearchFrom "James Kirk" --BookTitle "A Memoir" 
```

#### Libraries

In addition to some standard .NET stuff from Microsoft, this script uses some excellent libraries which made the entire thing tractable. 

* [MailKit](https://github.com/jstedfast/MailKit) - a simple to use IMAP client (among other things)
* [SmartReader](https://github.com/Strumenta/SmartReader) - a port of Mozilla's "readability" library which is the only reason I was able to write this using .NET instead of Node.
* [QuickEPUB](https://github.com/jonthysell/QuickEPUB) - A nice little library which wraps up all of the fiddly bits around generating epub files
* [Html Agility Pack](https://html-agility-pack.net) - a very full featured HTML parsing and manipulation library which is usually the first thing I reach for when I need to deal with HTML