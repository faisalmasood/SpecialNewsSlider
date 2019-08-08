# from https://www.c-sharpcorner.com/article/build-and-deploy-the-client-side-web-part-spfx-in-sharepoint-online/
gulp clean
gulp bundle --ship
gulp package-solution --ship
