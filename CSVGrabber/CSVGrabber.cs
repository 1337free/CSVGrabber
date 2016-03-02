using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;
using System.Net;
using LumenWorks.Framework.IO.Csv;
using System.IO;

namespace CSVGrabber
{
  public partial class CSVGrabber : ServiceBase
  {

    private readonly String URL_Location;
    private readonly String Save_Location;
    System.Timers.Timer _timer;
    DateTime _scheduleTime; 

    public CSVGrabber()
    {
      InitializeComponent();
      eventLog1 = new System.Diagnostics.EventLog();
      if (!System.Diagnostics.EventLog.SourceExists("CSVGrabber"))
      {
        System.Diagnostics.EventLog.CreateEventSource("CSVGrabber", "CSVGrabber Events");
      }
      eventLog1.Source = "CSVGrabber";
      eventLog1.Log = "CSVGrabber Events";
      RegistryKey SoftwareKey = Registry.LocalMachine.OpenSubKey("Software\\CSVGetter", false);

      try
      {
        
        URL_Location = (String)SoftwareKey.GetValue("URLLocation");
        Save_Location = (String)SoftwareKey.GetValue("FileLocation");
        _scheduleTime = DateTime.ParseExact((String)SoftwareKey.GetValue("Time"), "HHmm", System.Globalization.CultureInfo.InvariantCulture);
        SoftwareKey.Close();
      }
      catch (Exception)
      {
              
      }
      finally
      {
        if (URL_Location == null || Save_Location == null || _scheduleTime == null || URL_Location == String.Empty || Save_Location == String.Empty || _scheduleTime.Equals(new DateTime()))
        {
          eventLog1.WriteEntry("Registry information not found. Service Stopping.", EventLogEntryType.Information, 404);
          if (SoftwareKey != null)
          {
            SoftwareKey.Close();
          }
          ServiceController sc = new ServiceController("CSVGrabber");
          sc.Stop();
          sc.Close();
        }

      }



      


      _timer = new System.Timers.Timer();
      if (_scheduleTime < DateTime.Now) {
        _scheduleTime = _scheduleTime.AddDays(1);
      }


    }

    protected override void OnStart(string[] args)
    {
      eventLog1.WriteEntry("CSVGetter Service Started.");

      // For first time, set amount of seconds between current time and schedule time
      _timer.Enabled = true;
      _timer.Interval = _scheduleTime.Subtract(DateTime.Now).TotalSeconds * 1000;
      _timer.Elapsed += new System.Timers.ElapsedEventHandler(Timer_Elapsed);

    }

    protected void Timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
    {
      // 1. Process Schedule Task
      // ----------------------------------
      // Add code to Process your task here
      // ----------------------------------

      eventLog1.WriteEntry("Running according to scheduled time.", EventLogEntryType.Information, 1337);

      WebClient Client = new WebClient();
      try
      {
        Client.DownloadFile(URL_Location, System.IO.Path.GetTempPath().ToString() + "temp.csv");
      }
      catch (Exception)
      {
        
        eventLog1.WriteEntry("Error! The entered URL: \"" + URL_Location + "\" is not valid or you do not have sufficient permissions to access it.", EventLogEntryType.Information, 404);
        return;
      }
      
      Client.Dispose();

      StreamReader strmReader;
      StreamWriter strmWriter;
      CsvReader strmCSVReader;

      try
      {
        strmReader = File.OpenText(System.IO.Path.GetTempPath().ToString() + "temp.csv");
        strmWriter = new StreamWriter(Save_Location);
        strmCSVReader = new CsvReader(strmReader, true);

        foreach (String[] value in strmCSVReader.ToArray())
        {
          strmWriter.WriteLine(value[0]);
        }

        strmReader.Close();
        strmWriter.Close();
        strmCSVReader.Dispose();

        eventLog1.WriteEntry("CSV file was processed succesfully. Output should be located at " + Save_Location, EventLogEntryType.Information, 1);

      }
      catch (Exception)
      {

        eventLog1.WriteEntry("Error! Was unable to process CSV file, or invalid output location.", EventLogEntryType.Information, 505);
      }


      // 2. If tick for the first time, reset next run to every 24 hours
      if (_timer.Interval != 24 * 60 * 60 * 1000)
      {
        _timer.Interval = 24 * 60 * 60 * 1000;
      }
    }

    protected override void OnStop()
    {
      eventLog1.WriteEntry("CSVGrabber Service Stopped.");
    }
  }
}
