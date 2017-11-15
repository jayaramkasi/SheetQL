# SheetQL
This is a google sheets addon which helps Query mysql from Google sheets directly. According to Google documentation, " In Apps Script, the JDBC service supports Google Cloud SQL, MySQL, Microsoft SQL Server, and Oracle databases. "

You can deploy this as a standalone script or a bound script. As far as I can see, there is no particular difference in how you deploy to how the script runs. 

<b>How to deploy</b>

* Copy the code into a standalone google scripts project. It should have 3 files 
  - Code.gs
  - Sidebar.html
  - Connection.html

I use leonhartX's google chrome extension  https://chrome.google.com/webstore/detail/google-apps-script-github/lfjcgcmkmjjlieihflfhjopckgpelofo for github syncing.
* Deploy it as a sheets add on, within your domain. - Instructions provided <a href="https://developers.google.com/apps-script/add-ons/domain-wide">here</a>

Have fun!
