# NessusParser
A tool to create a vulnerability assessment document from a Nessus' csv file

<br>
<h2>Modules</h2>

In order to run the script you must install <code>python-docx</code> module.

Using the command: <code>pip3 install python-docx</code>.

<br>
<h2>Usage</h2>

<code>python3 NessusParser.py path/to/csv-files</code>

Where <code>path/to/csv-files</code> is a folder containig the csv-formatted result of a Nessus' scan, in order to create one or more of these files you should see [How to create a scan report in Nessus](https://docs.tenable.com/nessus/Content/CreateAScanReport.htm).

The script will do the following actions:
 - takes all the files in the folder "<code>path/to/csv-files</code>"
 - magically convert every single file into word format
    - if this is not possible, the program writes an <code>error.log</code> file containing the defective file(s) name.
 - insert them into a copy of the <code>template.docx</code>** file
 - asks you where to save the copy (the default name will be <code>"result.docx"</code>)

<br>

** If you want to use a different template file you must consider these aspects:

- The file "template.docx" must be into the same folder as the program.</b>
   - it must contain a string named <code>[VULNS HERE]</code> inside the document
- the "High", "Critical" and "Medium" Word styles must be present in the template document, if they are not, you must create them.

By default, the program generates the error file into its own folder.

<b>Please note:</b> the resulting document will not be pretty, if you want to make it presentable for a client you must edit it by yourself.
