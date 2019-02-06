# GE-Proficy-Historian
The Proficy Historian is a General Electric (GE) product for storing SCADA data. Similar to the more popular PI Historian by OSI Soft.

This project implements a python wrapper for the GE Proficy Historian SDK using COM Interop. The documentation for the Proficy Historian includes some VBA usage examples, but does not include any python code or wrapper code for abstracting the use of SDK objects.

The code is tested and working, but the library could do with a tidy up. Also, more development is needed to parse output XML files and get the dataset(s) out. There are also issues with the volume of data that can be returned (but this is probably more of a hardware/server memory resource limit). I solved these by breaking queries for large time frames into smaller more easy to manage chunks and joining the datasets with some post-processing.

## Perform a query
```python
import datetime.datetime
from Proficy import iHistorian

starttime = datetime(2010, 1, 1)
endtime = datetime(2011, 3, 31)

ConfigParams = 
{
'iHistorian' : [("ServerAddress", 'historian.domain'), 
                 ("Username", ''), 
                 ("Password", '')],
'SearchTags' : [("Tagmask", '*_AI_*'), 
                ("Tags", [])],
'Sampling' : [("SamplingMode", 'Calculated'), 
              ("CalculationMode", 'Average'), 
              ("Direction", 'Forward'), 
              ("NumberOfSamples", 0), 
              ("SamplingInterval", 30)],
'Timeframe' : [("StartTime", starttime), 
               ("EndTime", endtime)],
'Filtering' : [("FilterTag", ''),
               ("FilterComparisonMode", ''),
               ("FilterMode", ''),
               ("FilterComparisonValue", '')]
}

iHist = iHistorian(servername='myserver.domain', username='admin', password='')
iHist.throttle_queries(10e8, 600)

# Create a new data recordset
record = iHist.new_recordset("Data")

# Set constrains on the data record - sampling methodoligy
iHist.set_timeframe(recordset, starttime, endtime)
#iHist.set_sampling_from_parser(record, ConfigParams)
iHist.set_sampling_from_parser(record, p.ConfigParams)

# Set the SCADA tagname(s) to query
record.Criteria.Tags = ["SCADA1.TAG_1", "SCADA1.TAG_2"]

# Run the query
error = iHist.run_query(record)

# Get the data (this is very slow and runs on the server, so not ideal...
if not error:
  for tag in range(1, 1 + record.Tags.Count()):
    data = [] # Copys one column of data at a time
    for irecord in range(1, 1 + record.Item(tag)[0].Count()):
      iData = dataRecord.Item(1)[0].Item(irecord)
      data.append([iData.TimeStamp])
    if len(data) == 0:
      print '\nBad Tag. Skipping data."
else:
  print error
  
# Other ways to get the data that may be faster or more reliable
print record.XML()
record.Export(r"C:\SCADA CSV.csv", c.CSV)
```

### The GE Proficy Historian Excel plug-in may need to be loaded manually in Excel
```python
from win32com.client import constants as c, Dispatch

xlApp = Dispatch("Excel.Application")

def launchiHistorian(self):
	"""This is for loading the GE Proficy Historian Excel Add-in"""
	iHistorianPath = r"C:\Program Files (x86)\Microsoft Office\Office14\Library\iHistorian.xla"
	
	xlApp.DisplayAlerts = False
	assert os.path.exists(iHistorianPath) == True, "Error: There was a problem locating the iHistorian.xla Excel plug-in."
	xlApp.Workbooks.Open(iHistorianPath)
	xlApp.RegisterXLL(iHistorianPath)
	xlApp.Workbooks(iHistorianPath.split("\\")[-1]).RunAutoMacros = True
	xlApp.DisplayAlerts = True
 ```
