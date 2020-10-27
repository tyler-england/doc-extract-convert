# doc-extract-convert
Document contents -> Excel

Customers give us specification documents (URS) and my team must respond to the items in those documents (Accept/Refuse/Need Add'l Info). The documents may be received in DOC or PDF format, neither of which allows the Engineering team to easily respond to each individual requirement.

This started as a Python script which allowed me to extract the contents of a URS and populate an Excel workbook with each item on a separate line. But I ended up incorporating it into a Django project so that coworkers are able to access the functionality without needing to confer with me, and without having Python.

Now any user on our network can navigate to the site in a browser, upload a URS document, and receive the exported/properly-formatted Excel workbook.
