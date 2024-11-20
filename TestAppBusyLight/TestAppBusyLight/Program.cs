using Microsoft.Win32;
using System.Management;
using System.Security.Principal;
using System.IO.Ports;
using System.Text;
using System.Text.Json;
using static System.Net.WebRequestMethods;
//using System;
//using System.Net.Http;

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// define which LED types to update
const bool updateWledWlanJSON = true;
const bool updateWledSerialJSON = false;
const bool updateArduinoSerialPort = false;
const bool updateHomeAssistant = true;

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

//Arduino definitions
const string arduinoComPort = "COM14";
const int arduinoComPortSpeed = 9600;

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// Home Assistant definitions
const string homeAssistantUrl = "http://10.0.0.2:8123/";
const string homeAssistantStatesEndpoint = "/api/states";
const string homeAssistantEntity = "sensor.teams_status";
const string homeAssistantAPIKey = "xxxxxxxxx";

/*
Home assistan configuration.xml settings:

input_text:
  teams_status:
    name: Microsoft Teams status
    icon: mdi:microsoft-teams
sensor:
  - platform: template
    sensors:
      teams_status: 
        friendly_name: "Microsoft Teams status"
        value_template: "{{states('input_text.teams_status')}}"
        icon_template: "{{state_attr('input_text.teams_status','icon')}}"
        unique_id: sensor.teams_status
 */

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// WLED definitions
const string wledComPort = "COM7";
const int wledComPortSpeed = 115200;
const string wledJsonUrl = "http://10.0.0.5/json/";

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// Teams camera and micorphone registry parameters
const string appRegistryKey = "MSTeams_8wekyb3d8bbwe";
const string webcamPath = "webcam";
const string microphonePath = "microphone";
const string hive = "HKEY_USERS";
const string consentPath = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\CapabilityAccessManager\\ConsentStore\\";
const string stopValueName = "LastUsedTimeStop";


//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// begin of main program
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

var currentUser = WindowsIdentity.GetCurrent();
var webcamKeyPath = $"{currentUser.Owner.Value}\\{consentPath}{webcamPath}\\{appRegistryKey}\\";
var microphoneKeyPath = $"{currentUser.Owner.Value}\\{consentPath}{microphonePath}\\{appRegistryKey}\\";

string lastUsedTimeStopValueWebcam = Registry.GetValue($"{hive}\\{webcamKeyPath}", stopValueName, 0L).ToString();
string lastUsedTimeStopValueMicrophone = Registry.GetValue($"{hive}\\{microphoneKeyPath}", stopValueName, 0L).ToString();

bool webcamIsOn = false;
bool microphoneIsOn = false;

Console.Write("Startup lastUsedTimeStopValueWebcam: ");
Console.WriteLine(lastUsedTimeStopValueWebcam.ToString());
Console.Write("Startup lastUsedTimeStopValueMicrophone: ");
Console.WriteLine(lastUsedTimeStopValueMicrophone.ToString());

var dateAndTime = DateTime.Now;
Console.Write(dateAndTime.ToString("MM/dd/yyyy HH:mm:ss"));

if (lastUsedTimeStopValueWebcam is "0")
{
    Console.WriteLine(" Started app and found webcam to be on!");
    webcamIsOn = true;
    await UpdateLedStatus();
}
else

  if (lastUsedTimeStopValueMicrophone is "0")
{
    Console.WriteLine(" Started app and found microphone to be on!");
    microphoneIsOn = true;
    await UpdateLedStatus();
}
else
{
    Console.WriteLine(" Started app and found no call in progress!");
    await UpdateLedStatus();
}

SubscribeToRegistryEvents(stopValueName, webcamKeyPath, async (evt) => await WebcamLastUsedTimeStopHandler(evt));
SubscribeToRegistryEvents(stopValueName, microphoneKeyPath, async (evt) => await MicrophoneLastUsedTimeStopHandler(evt));

Console.WriteLine("Press the ENTER key to stop the program.");
Console.ReadLine();

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// end of main program
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

void SubscribeToRegistryEvents(string valueName, string path, Action<EventArrivedEventArgs> handler)
{
    var query = new WqlEventQuery($"SELECT * FROM RegistryValueChangeEvent WHERE Hive='{hive}' AND KeyPath='{path.Replace("\\", "\\\\")}' AND ValueName='{valueName}'");
    var watcher = new ManagementEventWatcher(query);
    watcher.EventArrived += (sender, args) => handler(args);
    watcher.Start();
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

async Task WebcamLastUsedTimeStopHandler(EventArrivedEventArgs args)
{
    var dateAndTime = DateTime.Now;
    Console.Write(dateAndTime.ToString("yyyy/MM/dd HH:mm:ss"));

    lastUsedTimeStopValueWebcam = Registry.GetValue($"{hive}\\{webcamKeyPath}", stopValueName, 0L).ToString();
    Console.Write(lastUsedTimeStopValueWebcam);

    if (lastUsedTimeStopValueWebcam is null or "0")
    {
        webcamIsOn = true;
        Console.WriteLine(": Webcam started");
    }
    else
    {
        webcamIsOn = false;
        Console.WriteLine(": Webcam stopped");
    }
    await UpdateLedStatus();
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

async Task MicrophoneLastUsedTimeStopHandler(EventArrivedEventArgs args)
{
    var dateAndTime = DateTime.Now;
    Console.Write(dateAndTime.ToString("yyyy/MM/dd HH:mm:ss"));

    lastUsedTimeStopValueMicrophone = Registry.GetValue($"{hive}\\{microphoneKeyPath}", stopValueName, 0L).ToString();
    Console.Write(lastUsedTimeStopValueMicrophone);

    if (lastUsedTimeStopValueMicrophone is null or "0")
    {
        microphoneIsOn = true;
        Console.WriteLine(": Microphone started");
    }
    else
    {
        microphoneIsOn = false;
        Console.WriteLine(": Microphone stopped");
    }
    await UpdateLedStatus();
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

async Task SendNetworkWledJSON(String jsonString, string jsonUrl)
{
    var jsonContent = new StringContent(jsonString, Encoding.UTF8, "application/json");
    HttpClient wledHttpJsonClient = new HttpClient();
    var jsonResponse = await wledHttpJsonClient.PostAsync(jsonUrl, jsonContent);

    if (jsonResponse.IsSuccessStatusCode)
    {
        await jsonResponse.Content.ReadAsStringAsync();
        Console.WriteLine("WLED JSON over network sent successfully.");
        return;
    }
    else
    {
        // Handle the error
        Console.Write("Unable to communicate with WLED over network using ");
        Console.WriteLine(wledJsonUrl);
        Console.WriteLine(jsonResponse.Content.ReadAsStringAsync());
        Console.WriteLine("HttpResponseMessage: ");
        Console.WriteLine(jsonResponse);
        return;
    }
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

async Task SendSerialPortWled(String jsonString)
{
    SerialPort _serialPortWled = new SerialPort($"{wledComPort}", wledComPortSpeed, Parity.None, 8, StopBits.One);
    _serialPortWled.Handshake = Handshake.None;
    _serialPortWled.Open();
    _serialPortWled.WriteLine(jsonString);
    _serialPortWled.Close();
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

async Task SendSerialPortArduino(String arduinoLedColour)
{
    SerialPort _serialPortArduino = new SerialPort($"{arduinoComPort}", arduinoComPortSpeed, Parity.None, 8, StopBits.One);
    _serialPortArduino.Handshake = Handshake.None;
    _serialPortArduino.Open();
    _serialPortArduino.WriteLine(arduinoLedColour);
    _serialPortArduino.Close();
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

async Task SendHomeAssistantStatus(string haStatus)
{
    var statusRequest = new SetStatusRequest { state = haStatus, attributes = new homeAssistantAttributes { friendly_name = "Microsoft Teams status", icon = "mdi:microsoft-teams" } };
    var jsonString = JsonSerializer.Serialize(statusRequest);

    Console.WriteLine(jsonString);

    var jsonContent = new StringContent(jsonString, Encoding.UTF8, "application/json");
    Console.WriteLine(jsonContent);

    using var homeAssistantClient = new HttpClient();
    homeAssistantClient.BaseAddress = new Uri(homeAssistantUrl);
    homeAssistantClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", homeAssistantAPIKey);

    var haJsonResponse = await homeAssistantClient.PostAsync($"{homeAssistantStatesEndpoint}/{homeAssistantEntity}", jsonContent);


    if (haJsonResponse.IsSuccessStatusCode)
    {
        await haJsonResponse.Content.ReadAsStringAsync();
        Console.WriteLine("Homeassistant over network sent successfully.");
        return;
    }
    else
    {
        // Handle the error
        Console.Write("Unable to communicate with Homeassistant over network using ");
        Console.WriteLine(homeAssistantUrl);
        Console.WriteLine(haJsonResponse.Content.ReadAsStringAsync());
        Console.WriteLine("HttpResponseMessage: ");
        Console.WriteLine(haJsonResponse);
        return;
    }
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

async Task UpdateLedStatus()
{
    if (updateArduinoSerialPort)
    {
        var arduinoLedStatusColour = webcamIsOn ? "R" : microphoneIsOn ? "r" : "g";
        await SendSerialPortArduino(arduinoLedStatusColour);
    }

    if (updateWledWlanJSON)
    {
        /* in WLED setup presets to reflect the effect you want:
        ps 1: green, effect solid color
        ps 2: red, effect solid color
        ps 3: red, effect blink
        */
        var wledJSON = webcamIsOn ? "{ps: 3}" : microphoneIsOn ? "{ps: 2}" : "{ps: 1}";
        await SendNetworkWledJSON(wledJSON, wledJsonUrl);
    }

    if (updateWledSerialJSON)
    {
        /* in WLED setup presets to reflect the effect you want:
        ps 1: green, effect solid color
        ps 2: red, effect solid color
        ps 3: red, effect blink
        */

        var wledJSON = webcamIsOn ? "{ps: 3}" : microphoneIsOn ? "{ps: 2}" : "{ps: 1}";
        await SendSerialPortWled(wledJSON);
    }

    if (updateHomeAssistant)
    {
        var status = webcamIsOn ? "VideoOn" : microphoneIsOn ? "InCall" : "Available";
        await SendHomeAssistantStatus(status);
    }
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

record SetStatusRequest
{
    public required string state { get; init; }
    public required homeAssistantAttributes attributes { get; init; }
}

record homeAssistantAttributes
{
    public required string friendly_name { get; init; }
    public string icon { get; init; }
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////