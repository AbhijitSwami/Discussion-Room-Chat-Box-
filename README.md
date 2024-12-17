## ChatBox## Discussion Room

rwbasechat is a simple chat webpart to include as webparts on SharePoint Online sites. It allows you to communicate through chat messages with your SharePoint audience.
To achieve this, this webpart uses a SharePoint list where the messages are stored. There is emoji support, but still a lot to do, especially on security and stability issues.


![image](https://github.com/user-attachments/assets/01c5c22a-66b3-49b2-91e3-8c3b5b0975ad)



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
