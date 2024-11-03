using Microsoft.Win32;
using System.Management;
using System.Security.Principal;

const string appRegistryKey = "MSTeams_8wekyb3d8bbwe";
const string webcamPath = "webcam";
const string microphonePath = "microphone";
const string hive = "HKEY_USERS";
var currentUser = WindowsIdentity.GetCurrent();
const string consentPath = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\CapabilityAccessManager\\ConsentStore\\";
var webcamKeyPath = $"{currentUser.Owner.Value}\\{consentPath}{webcamPath}\\{appRegistryKey}\\";
var microphoneKeyPath = $"{currentUser.Owner.Value}\\{consentPath}{microphonePath}\\{appRegistryKey}\\";
const string startValueName = "LastUsedTimeStart";
const string endValueName = "LastUsedTimeStop";

var lastUsedTimeStartValueWebcam = Registry.GetValue($"{hive}\\{webcamKeyPath}", startValueName, "0");
var lastUsedTimeStopValueWebcam =  Registry.GetValue($"{hive}\\{webcamKeyPath}", endValueName, "0");

var lastUsedTimeStartValueMicrophone = Registry.GetValue($"{hive}\\{microphoneKeyPath}", startValueName, 0l);
var lastUsedTimeStopValueMicrophone = Registry.GetValue($"{hive}\\{microphoneKeyPath}", endValueName, 0l);

var lastUsedTimeStartWebcam = DateTime.FromFileTime((long)lastUsedTimeStartValueWebcam);
var lastUsedTimeStopWebcam = DateTime.FromFileTime((long)lastUsedTimeStopValueWebcam);

var lastUsedTimeStartMicrophone = DateTime.FromFileTime((long)lastUsedTimeStartValueMicrophone);
var lastUsedTimeStopMicrophone = DateTime.FromFileTime((long)lastUsedTimeStopValueMicrophone);

bool webcamIsOn = false;
bool microphoneIsOn = false;

if(lastUsedTimeStartWebcam > lastUsedTimeStopWebcam)
{
    Console.WriteLine("Started app and found webcam to be on!");
    webcamIsOn = true;
}

if (lastUsedTimeStartMicrophone > lastUsedTimeStopMicrophone)
{
    Console.WriteLine("Started app and found microphone to be on!");
    microphoneIsOn = true;
}


SubscribeToRegistryEvents(startValueName, webcamKeyPath, WebcamLastUsedTimeStartHandler);
SubscribeToRegistryEvents(endValueName, webcamKeyPath, WebcamLastUsedTimeStopHandler);
SubscribeToRegistryEvents(startValueName, microphoneKeyPath, MicrophoneLastUsedTimeStartHandler);
SubscribeToRegistryEvents(endValueName, microphoneKeyPath, MicrophoneLastUsedTimeStopHandler);

Console.ReadLine();

void SubscribeToRegistryEvents(string valueName, string path, Action<EventArrivedEventArgs> handler)
{

    var query = new WqlEventQuery($"SELECT * FROM RegistryValueChangeEvent WHERE Hive='{hive}' AND KeyPath='{path.Replace("\\", "\\\\")}' AND ValueName='{valueName}'");
    var watcher = new ManagementEventWatcher(query);
    watcher.EventArrived += (sender, args) => handler(args);
    watcher.Start();
}

void WebcamLastUsedTimeStartHandler(EventArrivedEventArgs args)
{
    webcamIsOn = true;
    Console.WriteLine("Webcam started!");
}

void WebcamLastUsedTimeStopHandler(EventArrivedEventArgs args)
{
    if (webcamIsOn)
    {
        webcamIsOn = false;
        Console.WriteLine("Webcam stopped!");
    }
}

void MicrophoneLastUsedTimeStartHandler(EventArrivedEventArgs args)
{
    microphoneIsOn = true;
    Console.WriteLine("Microphone started!");
}

void MicrophoneLastUsedTimeStopHandler(EventArrivedEventArgs args)
{
    if (microphoneIsOn)
    {
        microphoneIsOn = false;
        Console.WriteLine("Microphone stopped!");
    }
}

