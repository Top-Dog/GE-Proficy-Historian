'''
Created on 10/11/2015

@author: SDO
@description:
A high level python wrapper for the GE Proficy Historian.
This handles connecting to the server and retrieving data (recordsests).
It allows for some query parameter tweaking, for example,
the max. sample count or the number of seconds until a timeout occurs.

'''

import sys, datetime, pythoncom
from sympy.physics.units import milliseconds

sys.path.append(r'C:\Python27\lib\site-packages\win32com')
from pywintypes import com_error
from win32com.client import constants as c
from win32com.client.gencache import EnsureDispatch

class iHistorian(object):
    def __init__(self, servername='myserver.domain', **kwargs):
        """
        @param servername: A string containing the historian server's URL or IP address
        @param username: A string containing the username. If one exists.
        @param password: A string containing the password. If one exists.
        """
        # Only create a class instanace and don't try connecting if no servername is supplied
        if servername:
            try:
                #DispatchEx 
                #=CreateObject("iHistorian_SDK.Server") or =New iHistorian_SDK.Server
                
                # -- Using comtypes libriary
                # from comtypes.client import CreateObject
                # Aquator = CreateObject("Aquator.Application")
                
                
                # Achieves early dispatch
                # http://timgolden.me.uk/python/win32_how_do_i/generate-a-static-com-proxy.html
                # http://stackoverflow.com/questions/170070/what-are-the-differences-between-using-the-new-keyword-and-calling-createobject
                self.ihApp = EnsureDispatch('iHistorian_SDK.Server') # Attach the server
                print "Connecting to: %s" % servername
                if not self._connect(servername, kwargs.get('Username'), kwargs.get('Password')):
                    print "Could not connect to the server at: %s" % servername
                    print self.ihApp.LastError
                    raw_input("Press 'Enter' to exit")
                    self.close()
                    exit()
            except TypeError:
                print "\nCannot dispatch iHistorian_SDK COM object."
                print "Check that ihSDK.dll and ihAPI40.dll are present and GE iHistorian is installed."
                print "Use regsvr32 to register the DLLs"
                print "cd 'C:\Python27\Lib\site-packages\win32com\client' select iHistorian and py code is automatically generated"
                raw_input("Press 'Enter' to exit")
                exit()
            
            print "SCADA server Connection successful! There are currently %d users connected to the server" % (self.num_of_connected_users())
        else:
            # Dispatch and instanciate the server SDK without connecting to a server - useful for building an index of ihistorian constants
            self.ihApp = EnsureDispatch('iHistorian_SDK.Server') # Attach the server
            
    def _connect(self, servername, username, password):
        """
        @return: True if there was a successful connection, False otherwise.
        """
        return self.ihApp.Connect(servername, username, password)[0]
        
    def _disconnect(self):
        """
        @return: The state of the sever after a call to disconnect it
        """
        return self.ihApp.Disconnect()
    
    def close(self):
        """
        Close and free-up memory
        """
        self._disconnect()
        #self.ihApp.Quit()
        self.ihApp = None
        del self.ihApp
        
    def throttle_queries(self, MaxSamplesPerTag=100000, MaximumQueryTime=60):
        """Restricts the max. number of samples per tag that the server can return
        (excludes raw and filtered data queries). Also sets the server's max. query 
        time (in seconds) before a connection is terminated."""
        self.ihApp.MaximumQueryIntervals = MaxSamplesPerTag
        self.ihApp.MaximumQueryTime = MaximumQueryTime 

    def new_recordset(self, RecordType):
        """
        Creates a new empty recordset object of the supplied type
        @return: A reference to the <type>recordset object
        """
        ValidRecordNames = ["Alarms", "Archives", "Collectors", "Data",
                            "Messages", "Tags"]
        assert RecordType in ValidRecordNames, "The supplied record: %s is not valid" % ValidRecordNames
        return getattr(self.ihApp, RecordType).NewRecordset()
        # to commit changes to a tagrecord use: tagRecord.WriteRecordset() 
        
    def clear_recordset(self, Recordset):
        """
        Clears the supplied recordset, so that the recorset object is completely blank
        """
        #Recordset.Clear()
        Recordset.Criteria.Clear()
        Recordset.Fields.Clear()   
        
    def set_query_fields(self, recordset, fieldlist):
        """Allows you to choose what parameters you want the query 
        to return. The parameters are supplied as a list of strings,
        for example, ["Value", "TimeStamp", "Comments"]
        @return: True for success, False otherwise."""
        return recordset.SetFields(fieldlist)        
        
    def quality_sample(self, dataRecord, maxValue, minValue, defaultValue):
        """
        @param dataRecord: An iHistorian data record with the query attributes added and the query data returned
        """
        for tag in range(1, 1 + dataRecord.Tags.Count()):
            for record in range(1, 1 + dataRecord.Item(tag)[0].Count()):
                iData = dataRecord.Item(tag)[0].Item(record)
                if iData.DataQuality == 1:
                    print "Good"
                    print iData.Value # actual value of the tag
                    print iData.TimeStamp # in pytime
                    
                    k = iData.Comments.Count
                    print iData.Comments(k).Comment
                    # To get equispaced samples; specify either NumberOfSamples or SamplingInterval (int type milliseconds)
                elif iData.DataQuality == 2:
                    print "Bad"
                elif iData.DataQuality == 3:
                    print "Unknown"
                else:
                    print "Data error"
                    
    def build_query_from_parser(self, DataRecordset, ConfigParams):
        """
        @param DataRecordset: A datarecordset object created by Proficy Historian for a query
        @param ConfigParams: An instance of the parser class' ConfigParams, which contains the query parameters
        Builds an iHistorian query from a datarecordset object using the parameters provided in the
        [parsed] xlsx configuration file.
        """
        f = DataRecordset.Criteria
        
        # Loop through each keyword in the dictionary
        for ConfigClass in ConfigParams:
            # Loop through each tuple in the key:value pair (each tuple represents a unique parameter)
            for param in ConfigParams.get(ConfigClass):
                # set values only if they are criteria for the datarecordset
                # Param[0] is a string that is the label of the paramter
                # Param[1] is variable (string, list, int) that is the value of the query param (set in FileIO2.read_config)
                if ConfigClass is "Sampling" and param[1] is not None:
                    setattr(f, param[0].replace(' ', ''), param[1]) # Create a constraint in the query
                
                # Set the SCADA tags to retrieve information for (from the Substation objects)
                if ConfigClass is "Substations" and param[1] is not None:
                    #setattr(f, "Tags", param[1]) <-- this only puts the substation names in not their tags
                    setattr(f, "Tags", ['SCADA.AHL_AI_3BA.F_CV','SCADA.AHL_AI_32BV.F_CV'])
                if ConfigClass is "Timeframe2" and param[1] is not None:
                    setattr(f, param[0].replace(' ', ''), param[1])
          
        # Set the filters (if there are any set)
        #if f.FilterTag:
        #    f.FilterTagSet = True
        #if f.FilterComparisonMode:
        #    f.FilterComparisonModeSet = True
        #if f.FilterMode:
        #    f.FilterModeSet = True
        
    def set_sampling_from_parser(self, DataRecordset, ConfigParams):
        """
        Sets up the sampling method(s) to be used with this particular queryset.
        
        @param DataRecordset: A datarecordset object created by Proficy Historian for a query
        @param ConfigParams: An instance of the parser class' ConfigParams, which contains the query parameters
        Builds an iHistorian query from a datarecordset object using the parameters provided in the
        [parsed] xlsx configuration file.
        """
        f = DataRecordset.Criteria
        
        # Loop through each keyword in the dictionary
        for ConfigClass in ConfigParams:
            # Loop through each tuple in the key:value pair (each tuple represents a unique parameter)
            for param in ConfigParams.get(ConfigClass):
                # set values only if they are criteria for the datarecordset
                # Param[0] is a string that is the label of the paramter
                # Param[1] is variable (string, list, int) that is the value of the query param (set in FileIO2.read_config)
                if ConfigClass is "Sampling" and param[1] is not None:
                    setattr(f, param[0].replace(' ', ''), param[1]) # Create a constraint in the query
    
    def getDateTime(self, pyDate):
        """
        @param pyDate: A date time object from COM or Excel
        @return: a conventional python datetime object
        """
        return datetime.datetime(pyDate.year, pyDate.month, pyDate.day,
                         pyDate.hour, pyDate.minute, pyDate. second,
                         pyDate.msec / 1000)
        
    def set_timeframe_from_parser(self, DataRecordset, ConfigParams):
        """
        Sets up the sampling method(s) to be used with this particular queryset.
        
        @param DataRecordset: A datarecordset object created by Proficy Historian for a query
        @param ConfigParams: An instance of the parser class' ConfigParams, which contains the query parameters
        Builds an iHistorian query from a datarecordset object using the parameters provided in the
        [parsed] xlsx configuration file.
        """
        f = DataRecordset.Criteria
        dates = []
        
        # Loop through each keyword in the dictionary
        for ConfigClass in ConfigParams:
            # Loop through each tuple in the key:value pair (each tuple represents a unique parameter)
            for param in ConfigParams.get(ConfigClass):
                # set values only if they are criteria for the datarecordset
                # Param[0] is a string that is the label of the paramter
                # Param[1] is variable (string, list, int) that is the value of the query param (set in FileIO2.read_config)
                if ConfigClass is "Timeframe1" and param[1] is not None:
                    dates.append(param[1]) # Append the pydate object
                if ConfigClass is "Timeframe2" and param[1] is not None:
                    dates.append(param[1]) # Append the pydate object
                
                #if ConfigClass is "Timeframe2" and param[1] is not None:
                #    setattr(f, param[0].replace(' ', ''), param[1])
        setattr(f, "StartTime", min(dates))
        setattr(f, "EndTime", max(dates))
        
    def set_timeframe(self, recordset, starttime, endtime):
        """Sets the start and end times with either pytime or datetime objects"""
        assert starttime < endtime, "Error: start time is greater than end time."
        recordset.Criteria.StartTime = starttime
        recordset.Criteria.EndTime = endtime      
    
    def run_part_query(self, timestart, ConfigParams, maxsamples=1000):
        """Limits the return of a query to 1000 samples per tag. 
        And steps through the required number of samples or date range
        to complete the initial query. """
        # Handle setting the timeframe of the query
        timedelta = datetime.timedelta(milliseconds=ConfigParams.get("Sampling")[4][1])
        timeend = maxsamples * timedelta
        
        return timeend + timedelta # this will be the new timestart
    
    def run_query(self, recordset):
        """
        @param recordset: An iHistorian (any type) recordset object with the query attributes added
        @return: Boolean. If a query was successful (True) or not (False)
        """
        recordset.Fields.Clear()
        recordset.Fields.AllFields()
        if recordset.QueryRecordset():
            return u''
        else:
            return recordset.LastError
        
    def export_record(self, recordset, path):
        """Creates an output file for the entire recordset.
        Intended for use with datarecordsets, as it is much 
        faster than reading individual values from a COM object.  
        Valid options for type are c.CSV, c.XML, c.Report
        @return: True for success, False otherwise."""
        filetype = 0 # Export will default to CSV type
        if path.endswith(".csv"):
            filetype = c.CSV
        elif path.endswith(".xml"):
            filetype = c.XML # This is the same as recordset.XML(), except you don't have to write an intermediate file, all in RAM
        elif path.endswith(".RPT"):
            filetype = c.Report
        return recordset.Export(path, filetype)[0]
    
    def export_execute_query(self, recordset):
        # recordset.XML([XMLHeader], [StartIndex], [EndIndex])
        # XMLHeader = XML to include before the TagRecordset XML 
        # StartIndex = Index of first tag to include in XML 
        # EndIndex = Index of last tag to include in XML 
        return recordset.XML()[0]
    
    # Utility functions below
    def num_of_tags(self):
        """
        @return: The (total) number of SCADA tags currently configured on the server.
        The number returned is less than or equal to the number of licensed tags.
        """
        return self.ihApp.ActualTags
    
    def num_of_connected_users(self):
        """
        @return: The number of users connected to the Historian server.
        The number returned is less than or equal to the number of licensed users.
        """
        return self.ihApp.ActualUsers
    
    
    
    # build query from inputs
    #def build_query(self, dataRecord, **kwargs):
    #    """
    #    Builds a query for iHistorian
    #    @param dataRecord: A reference to an empty iHistorian data record set object (to be populated)
    #    @param kwargs: All the required keys to build the query (e.g. Direction=c.Forward)
    #    All time argument values must be given using pytime as opposed to datetime
    #    """
    #    # Check that we have the minimum required keys to run a query
    #    requiredKeys = {"StartTime", "EndTime", "SamplingMode", "Direction"}
    #    suppliedKeys = set(kwargs.keys())
    #    if len(requiredKeys & suppliedKeys) < len(requiredKeys):
    #        print "Ensure that minimum required number/types of keys are supplied"
    #        exit()
    #    
    #    s = dataRecord.Criteria
    #    for name, value in kwargs.items():
    #        try:
    #            defaultValue = getattr(s, name)
    #        except AttributeError:
    #            print "%s is not a recognised variable name" % name
    #            exit()
    #        try:
    #            setattr(s, name, value)
    #        except:
    #            print "%s is not a valid value. Check variable type." % value
    #            exit()
    
    def build_query_from_parser_OLD(self, DataRecordset, ParseInstance):
        """
        @param DataRecordset: A datarecordset object created by Proficy Historian for a query
        @param ParseInstance: An instance of the parse class, which contains the query parameters
        Builds an iHistorian query from a datarecordset object using the parameters provided in the
        [parsed] xlsx configuration file.
        """
        f = DataRecordset.Criteria
        
        for ConfigClass in ParseInstance.ConfigParams:
            for param in ParseInstance.ConfigParams.get(ConfigClass):
                # set values only if they are criteria for the datarecordset
                if ConfigClass is not "iHistorian" and param[1] is not None:
                    setattr(ParseInstance, param[0], param[1]) # Create a local class variable
                    setattr(f, param[0], param[1]) # Create a constraint in the query
                elif ConfigClass is not "iHistorian":
                    setattr(ParseInstance, param[0], param[1])
                    
        # Set the filters (if there are any set)
        if f.FilterTag:
            f.FilterTagSet = True
        if f.FilterComparisonMode:
            f.FilterComparisonModeSet = True
        if f.FilterMode:
            f.FilterModeSet = True
    



# Classes needed for a DataRecordset query
class query_data_recordset():
    # Criteria Property (DataRecordset Object)
    def __init__(self, dataRecord, **kwargs): # TimeFrame, Tags, SamplingMode, CalculationMode, FilteringCriteria
        f = dataRecord.Criteria
        
        self.SamplingMode = getattr(f, kwargs.get("SamplingMode", c.Calculation))
        if self.SamplingMode == c.Calculation:
            self.CalculationMode = kwargs.get("CalculationMode", c.Average)
            f.CalculationMode = self.CalculationMode
        if self.SamplingMode == c.RawByNumber:
            self.Direction = kwargs.get("Direction", c.Forward)
            f.Direction = self.Direction
        
        # Sample Intervals (let the SDK decide which to use if both modes are non-zero)
        self.NumberOfSamples = kwargs.get("NumberOfSamples", 0)
        f.NumberOfSamples = self.NumberOfSamples
        self.SamplingInterval = kwargs.get("SamplingInterval", datetime.timedelta()).total_seconds() * 1000
        f.SamplingInterval = self.SamplingInterval
        
        # Set the start and end time/dates
        DefaultStartTime = f.StartTime 
        self.StartTime = kwargs.get("StartTime", DefaultStartTime)
        f.startTime = self.StartTime 
        DefaultEndTime = f.EndTime 
        self.EndTime = kwargs.get("EndTime", DefaultEndTime)
        f.EndTime = self.EndTime
        
        # Set the tags or the tagmask
        f.Tagmask = kwargs.get("Tagmask", "")
        
        # Set the filters (if there are any set)
        self.FilterTag = kwargs.get("FilterTag", "")
        if self.FilterTag != "":
            f.FilterTagSet = True
        self.FilterComparisonMode = kwargs.get("FilterComparisonMode", 0)
        if self.FilterComparisonMode != 0:
            f.FilterComparisonModeSet = True
        self.FilterMode = kwargs.get("FilterMode", 0)
        if self.FilterMode != 0:
            f.FilterModeSet = True
        self.FilterComparisonValue = kwargs.get("FilterComparisonValue", "")
            
            
    def query_enable_filter(self):
        pass
