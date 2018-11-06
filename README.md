## sp-fx-extn

This is where you include your WebPart documentation.

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

### Build options

gulp clean - TODO
gulp test - TODO
gulp serve - TODO
gulp bundle - TODO
gulp package-solution - TODO


extension deployment command:
=====================================
create extention:
++++++++++++++++++++++
check if everything installed correctly: dir %appdata%\npm\*.cmd /b
sharepoint generator: %appdata%\npm\yo @microsoft/sharepoint
Replace PageUrl in server.json
+++++++++++++++++++++++++++
Deployment cmd:
+++++++++++++++++++++++++++
%appdata%\npm\gulp trust-dev-cert
%appdata%\npm\gulp serve

