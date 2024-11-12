using Microsoft.Win32;
using System.Management;
using System.Security.Principal;
using System.IO.Ports;
using System.Text;
using System.Text.Json;

// define which LED types to update
const bool updateWledWlanJSON = true;
const bool updateWledSerialJSON = false;
const bool updateArduinoSerialPort = false;
const bool updateHomeAssistant = false;

//Arduino definitions
const string arduinoComPort = "COM14";
const int arduinoComPortSpeed = 9600;

// Home Assistant definitions
const string homeAssistantUrl = "https://10.30.0.20:8443/api";
const string homeAssistantStatesEndpoint = "/states/";
const string homeAssistantEntity = "entity_id";
const string homeAssistantAPIKey = "";

// WLED definitions
const string wledComPort = "COM7";
const int wledComPortSpeed = 115200;
const string wledJsonUrl = "http://192.168.178.169/json/";

// serial port definitions
SerialPort _serialPortArduino = new SerialPort($"{arduinoComPort}", arduinoComPortSpeed, Parity.None, 8, StopBits.One);
_serialPortArduino.Handshake = Handshake.None;

SerialPort _serialPortWled = new SerialPort($"{wledComPort}", wledComPortSpeed, Parity.None, 8, StopBits.One);
_serialPortWled.Handshake = Handshake.None;

// Team camera and micorphone registry parameters
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

Console.Write(lastUsedTimeStopValueWebcam.ToString());
Console.Write(lastUsedTimeStopValueMicrophone.ToString());

if (lastUsedTimeStopValueWebcam is "0")
{
    var dateAndTime = DateTime.Now;
    Console.Write(dateAndTime.ToString("MM/dd/yyyy HH:mm:ss"));
    Console.WriteLine("Started app and found webcam to be on!");
    webcamIsOn = true;
    await UpdateLedStatus();
}

if (lastUsedTimeStopValueMicrophone is "0")
{
    var dateAndTime = DateTime.Now;
    Console.Write(dateAndTime.ToString("MM/dd/yyyy HH:mm:ss"));
    Console.WriteLine("Started app and found microphone to be on!");
    microphoneIsOn = true;
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

async Task SendJSON(String jsonString, string jsonUrl)
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

async Task SendHA(string haStatus)
{
    var statusRequest = new SetStatusRequest { Status = haStatus, Attributes = new Attributes { FriendlyName = $"{haStatus} but friendly" } };

    var jsonString = JsonSerializer.Serialize(statusRequest);
    var jsonContent = new StringContent(jsonString, Encoding.UTF8, "application/json");
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
        Console.WriteLine(wledJsonUrl);
        Console.WriteLine(haJsonResponse.Content.ReadAsStringAsync());
        Console.WriteLine("HttpResponseMessage: ");
        Console.WriteLine(haJsonResponse);
        return;
    }
}


async Task UpdateLedStatus()
{
    if (updateArduinoSerialPort)
    {
        if (webcamIsOn)
        {
            _serialPortArduino.Open();
            _serialPortArduino.Write("R");
            _serialPortArduino.Close();
        }
        else
        {
            if (microphoneIsOn)
            {
                _serialPortArduino.Open();
                _serialPortArduino.Write("r");
                _serialPortArduino.Close();

            }
            else    // webcam and microphone off
            {
                _serialPortArduino.Open();
                _serialPortArduino.Write("g");
                _serialPortArduino.Close();
            }
        }
    }

    if (updateWledWlanJSON)
    {
        /* in WLED setup presets to reflect the effect you want:
        ps 1: green, effect solid color
        ps 2: red, effect solid color
        ps 3: red, effect blink
        */
        if (webcamIsOn)
        {
            //Send JSON load preset 3, red blink
            string wledJSON = "{ps: 3}";
            await SendJSON(wledJSON, wledJsonUrl);
        }
        else
        {
            if (microphoneIsOn)
            {
                //Send JSON load preset 2, red solid
                string wledJSON = "{ps: 2}";
                await SendJSON(wledJSON, wledJsonUrl);
            }
            else    // webcam and microphone off
            {
                //Send JSON load preset 1, green solid
                string wledJSON = "{ps: 1}";
                await SendJSON(wledJSON, wledJsonUrl);
            }
        }
    }

    if (updateWledSerialJSON)
    {
        /* in WLED setup presets to reflect the effect you want:
        ps 1: green, effect solid color
        ps 2: red, effect solid color
        ps 3: red, effect blink
        */
        if (webcamIsOn)
        {
            //Send JSON load preset 3, red blink
            _serialPortWled.Open();
            //_serialPortWled.Write("{\"v\":true}");  //enable JSON over serial
            string wledJSON = "{\"ps\": 3}";
            _serialPortWled.WriteLine(wledJSON);
            _serialPortWled.Close();
        }
        else
        {
            if (microphoneIsOn)
            {
                //Send JSON load preset 2, red solid
                _serialPortWled.Open();
                //_serialPortWled.Write("{\"v\":true}");  //enable JSON over serial
                string wledJSON = "{\"ps\": 2}";
                _serialPortWled.WriteLine(wledJSON);
                _serialPortWled.Close();
            }
            else    // webcam and microphone off
            {
                //Send JSON load preset 1, green solid
                _serialPortWled.Open();
                //_serialPortWled.Write("{\"v\":true}");  //enable JSON over serial
                string wledJSON = "{\"ps\": 1}";
                _serialPortWled.WriteLine(wledJSON);
                _serialPortWled.Close();
            }
        }
    }

    if (updateHomeAssistant)
    {
        var status = webcamIsOn ? "VideoOn" : microphoneIsOn ? "InCall" : "Available";
        await SendHA(status);
    }
}


record SetStatusRequest
{
    public required string Status { get; init; }
    public required Attributes Attributes { get; init; }
}

record Attributes
{
    public required string FriendlyName { get; init; }
    public string Icon => "mdi:microsoft-teams";

}
