See http://jxls.sf.net/ for online version

Jxls v2.4.7 is released!
=========================

Changes
-----------

jxls-2.4.7
 
* fixed [issue#125 Allow passing null image to ImageCommand](https://bitbucket.org/leonate/jxls/issues/125/image-command-customization)
* fixed [issue#124 Optimizing XlsArea:findCommandsForVerticalShift recursive method](https://bitbucket.org/leonate/jxls/issues/124/optimizing-xlsarea)
* fixed [issue#122 Wrong cell replacement](https://bitbucket.org/leonate/jxls/issues/122/wrong-cell-replacement)

jxls-poi-1.0.16

* support for Apache POI 4.0.0

jxls-reader-2.0.5

* support for Apache POI 4.0.0
* reverted the fix for "XLSBlockReader startRow/endRow not work" merged earlier in [PR#5](https://bitbucket.org/leonate/jxls-reader/pull-requests/5/fixed-for-xlsblockreader-startrow-endrow/diff) 


See description of previous releases at [Version History](changes.html)

The latest component versions are

* org.jxls:jxls:2.4.7

* org.jxls:jxls-poi:1.0.16

* org.jxls:jxls-jexcel:1.0.7

* org.jxls:jxls-reader:2.0.5


Introduction
------------
Jxls is a small Java library to make generation of Excel reports easy.
Jxls uses a special markup in Excel templates to define output formatting and data layout.

Excel generation is required in many Java applications that have some kind of reporting functionality.

Java has great open-source and commercial libraries for creating Excel files (of open source ones worth mentioning are [Apache POI](https://poi.apache.org/)
and [Java Excel API](http://jexcelapi.sourceforge.net/).

Those libraries are quite low-level in a sense that they require you to write a lot of Java code even to create simple Excel files.

Usually you have to manually set each cell formatting and data for the spreadsheet.
Depending on the complexity of the report layout and data formatting the Java code can become quite complex and difficult to debug and maintain.
In addition not all Excel features are supported and can be manipulated with API(for example macros or graphs).
The suggested workaround for unsupported features is to create the object manually in an Excel template  and fill the template with data after that.

Jxls takes this approach to a higher level. All you need to do when working with Jxls is just to define all your report formatting and data layout in an Excel template and run Jxls engine
 providing it with the data to fill in the template. The only code you need to write in the most cases is a simple invocation of Jxls engine with proper configuration.

Features
--------
* XML and binary Excel format output (depends on underlying low-level Java-to-Excel implementation)
* Java collections output by rows and by columns
* Conditional output
* [Expression language](reference/expression_language.html) in report definition markup 
* [Multiple sheets output](reference/multi_sheets.html)
* Native Excel formulas
* Parameterized formulas
* Grouping support
* Merged cells support
* Area listeners to adjust excel generation
* Excel comments mark-up for command definition
* XML mark-up for command definition
* Custom Command definition
