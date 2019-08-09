# Special News Slider
Updated SPFX40 news slider that pulls data from multiple source SharePoint webs and lists.   Reads data from SPView for user defined query support.  Based on https://github.com/OlivierCC/spfx-40-fantastics/wiki/News-Slider


# Web Part Type
SharePoint Framework

# SharePoint Environment Compatibility
<table>
    <tr>
        <th>SharePoint Online</th>
        <th>SharePoint 2013</th>
        <th>SharePoint 2016</th>
        <th>SharePoint 2019</th>
        <th>Classic</th>
        <th>Modern</th>
    </tr>
    <tr>
        <td style="background-color:green;color:white;">Yes</td>
        <td style="background-color:red;color:white;">No</td>
        <td style="background-color:green;color:white;">Yes</td>
        <td style="background-color:green;color:white;">Yes</td>
        <td style="background-color:green;color:white;">Yes</td>
        <td style="background-color:green;color:white;">Yes</td>
    </tr>
</table>


# File Details

## Config

### package-solution.js
SPFx solution configuration. Allows us to control the name of the solution as seen in SharePoint, the file name of the output .sppkg, as well as the version number of the solution.


# Screenshots


# Testing in a remote Office 365 developer site

* Open serve.json file and update the pageUrl value in default configuration node to your dev site workbench URL.</li>

```bash
gulp clean
gulp build
gulp bundle
gulp serve
```


### Development


| Command            | Purpose |
|-------------------------|----------------------------------------------------------------------------------------|
| `gulp clean`            | Remove generated items; sometimes helpful if things don't seem to be behaving properly ||
| `gulp build`            | Builds the solution                                                                    |
| `gulp bundle`           | Bundles resources together prior to packaging                                          |
| `gulp package-solution` | Packages the solution                                                                  |
| `gulp serve`            | Test locally in the SharePoint Workbench                                               |
### Production


| Command            | Purpose |
|-------------------------|----------------------------------------------------------------------------------------|
| `gulp clean`            | Remove generated items; sometimes helpful if things don't seem to be behaving properly ||
| `gulp build --ship`            | Builds the solution                                                                    |
| `gulp bundle --ship`           | Bundles resources together prior to packaging                                          |
| `gulp package-solution --ship` | Packages the solution                                                                  |

## Supported Languages
- English


