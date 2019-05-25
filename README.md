## react-teams-list

This `SPFX webpart` which can also be used as `Teams tab` lists all the teams in the tenant.
It will display if a user is already part of a `Microsoft Teams Team` and if not provide a `join` link

### In Teams

![teams](./images/teams-list-teams.png)

### In SharePoint Site

![teams](./images/teams-tab-sp.png)


## Used SharePoint Framework Version 
![1.8.1](https://img.shields.io/badge/version-1.8.1-green.svg)
`plusBeta`


Version|Date|Comments
-------|----|--------
1.0|May 6, 2019|Initial release

Solution|Author(s)
--------|---------
react-teams-list|Rabia Williams


### This webpart requires Web API Permissions for 

Make sure you make this an isolated webpart, for ease of demonstration this webpart is currently not an isolated webpart and uses existing permissions granted

```
{
"resource": "Microsoft Graph",
"scope": "User.Read.All"
},
{
"resource": "Microsoft Graph",
"scope": "User.ReadBasic.All"
},{
"resource": "Microsoft Graph",
"scope": "Groups.Read.All"
}
```


![permission](./images/permission.png)

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

## Disclaimer
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Package the SPFX webpart
- Copy the package `react-teams-list` to the appcatalog
- Copy the zip file in `teams` folder to the appcatalog
- Upload the zip file also to the `Teams` app, via `side loading`
- Add the app into a channel
