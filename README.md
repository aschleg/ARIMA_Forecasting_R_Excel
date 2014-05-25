## ARIMA Forecasting with Excel

### Getting Started

Before using the sheet, you must download R and RExcel from the <a href="http://rcom.univie.ac.at">Statconn website.</a> If you already have R installed, you can just download RExcel. If you don't have R installed, you can download RAndFriends which contains the latest version of R and RExcel. Please Note, RExcel only works on 32bit Excel for its non-commercial license. If you have 64bit Excel installed, you will have to get a commercial license from Statconn.

RExcel now comes with an auto-activation program that will connect it with Excel; however, I'll keep the process for manually installing just in case.

#### Manually installing RExcel

To install RExcel and the other packages to make R work in Excel, first open R as an Administrator by right-clicking on the .exe.

![alt text](http://www.aaronschlegel.com/wp-content/uploads/2013/05/r_admin.png "Run R as Administrator")

In the R console, install RExcel by typing the following statements:

```
library(RExcelInstaller)

installRExcel()
```

This will install RExcel on your machine.

The next step is to install rcom, which is another package from Statconn for the RExcel package. To install this, type the following commands. This will also automatically install rscproxy as of R version 2.8.0.

```
library(rcom)

installstatconnDCOM()

comRegisterServer()
```

With these packages installed, you can move onto to setting the connection between R and Excel.

Although not necessary to the installation, a handy package to download is Rcmdr, developed by John Fox. This creates R menus that can become menus in Excel. This comes by default with the RAndFriends installation, and makes several R commands available in Excel.

Type the following commands into R to install Rcmdr.

```
library(Rcmdr)

installRcmdr()
```

Now that RExcel and its dependencies are installed, we can create the link to R and Excel.

Note in recent versions of RExcel this connection is made with a simple double-click of the provided .bat file "ActivateRExcel2010", so you should only need to follow these steps if you manually installed R and RExcel or if for some reason the connection isn't made during the RAndFriends installation.

#### Create the Connection Between R and Excel

Open a new book in Excel and navigate to the options screen.

![alt text](http://www.aaronschlegel.com/wp-content/uploads/2013/05/exceloptions.png "Go to the Options panel in Excel")

Click Options and then Add-Ins. You should see a list of all the active and inactive add-ins you currently have. Click the 'Go' button at the bottom.

![alt text](http://www.aaronschlegel.com/wp-content/uploads/2013/05/addinscreen.png "List of available Add-ins")

On the Add-Ins dialog box, you will see all the add-in references you have made. Click on Browse.

![alt text](http://www.aaronschlegel.com/wp-content/uploads/2013/05/addinscreen.png "Click browse to find the RExcel Add-in")

Navigate to the RExcel folder, usually located in C:Program FilesRExcelxls or something similar. Find the RExcel.xla add-in and click it.

The next step is to create a reference in order for macros using R to work properly. In your Excel doc, enter Alt + F11. This will open Excel's VBA editor. Go to Tools -> References, and find the RExcel reference, 'RExcelVBAlib'. RExcel should now be ready to use!

![alt text](http://www.aaronschlegel.com/wp-content/uploads/2013/05/vbaeditor.png "References available in the VBA Editor")

### Using the Excel Sheet

Now that R and RExcel are properly configured, it's time to do some forecasting!

Open the forecasting sheet and click 'Load Server'. This is to start the RCom server and also load the necessary functions to do the forecasting. A dialog box will open. Select the 'functions.R' file included with the sheet. This file contains the functions the forecasting tool uses. Most of the functions contained were developed by <a href="http://www.stat.pitt.edu/stoffer/tsa3/">Professor Stoffer at the University of Pittsburgh</a>. They extend the capabilities of R and give us some nice diagnostic graphs along with our forecasting output. There is also a function to automatically determine the best fitting parameters of the ARIMA model.

Once the server is loaded, enter your data into the Data column. Select the range of the data, right-click and select 'Name Range'. Name the range as 'Data'.

![alt text](http://www.aaronschlegel.com/wp-content/uploads/2013/05/namerange.png "Name the range 'Data'")

Next, set the frequency of your data in Cell C6. The frequency refers to the time periods of your data. If it is weekly, the frequency would be 7. Monthly would be 12, while quarterly would be 4, and so on.

Enter the periods ahead to forecast. Note that ARIMA models become quite inaccurate after several successive frequency predictions. A good rule of thumb is not to exceed 30 steps as anything past that could be rather unreliable. This does depend on the size of your data set as well. If you have limited data available, it is recommended to choose a smaller steps ahead number.

After entering your data, naming it, and setting the desired frequency and steps ahead to forecast, click Run. It may take a while for the forecasting to process.

Once it's completed, you will get predicted values out to the number you specified, the standard error of the results, and two charts. The left is the predicted values plotted with the data, while the right contains handy diagnostics featuring standardized residuals, the autocorrelation of the residuals, a gg plot of the residuals and a Ljung-Box statistics graph to determine if the model is well fitted.

I won't get into too much detail on how you look for a well fitted model, but on the ACF graph you don't want any (or a lot) of the lag spikes crossing over the dotted blue line. On the gg plot, the more circles that go through the line, the more normalized and better fitted the model is. For larger datasets this might cross a lot of circles. Lastly, the Ljung-Box test is an article in itself; however, the more circles that are above the dotted blue line, the better the model is.

If the diagnostics result doesn't look good, you might try adding more data or starting at a different point closer to the range you want forecast.

You can easily clear the generated results by clicking the 'Clear Forecasted Values' buttons.

And that's it! Currently, the date column doesn't do anything other than for your reference, but it's not necessary for the tool. If I find time, I'll go back and add that so the displayed graph shows the correct time. You also might receive an error when running the forecast. This is usually due to the function that finds the best parameters is unable to determine the proper order. You can follow the above steps to try and arrange your data better for the function to work.

I also wrote an explanation on ARIMA and the mathematics behind it <a href="http://www.aaronschlegel.com/arima-forecasting-an-introduction-and-implementation-with-excel/">here</a>.

Big thanks to <a href="http://www.stat.pitt.edu/stoffer/tsa3/">Professor Stoffer</a> for making his functions available.

