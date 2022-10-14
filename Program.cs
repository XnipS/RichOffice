using DiscordRPC;
using DiscordRPC.Logging;
using System.Diagnostics;
using Newtonsoft.Json;

// Welcome text
Console.WriteLine("Welcome to: \r\n  _____  _      _        ____   __  __ _          \r\n |  __ \\(_)    | |      / __ \\ / _|/ _(_)         \r\n | |__) |_  ___| |__   | |  | | |_| |_ _  ___ ___ \r\n |  _  /| |/ __| '_ \\  | |  | |  _|  _| |/ __/ _ \\\r\n | | \\ \\| | (__| | | | | |__| | | | | | | (_|  __/\r\n |_|  \\_\\_|\\___|_| |_|  \\____/|_| |_| |_|\\___\\___|\r\n\nBy XnipS");

// Variables
bool appfound = false;
DiscordRpcClient client = null;
string noun = "idle";
string input = "";
Config configuration = null;

// Load Json
try
{
	using(StreamReader r = new StreamReader("config.json"))
	{
		string json = r.ReadToEnd();
		configuration = JsonConvert.DeserializeObject<Config>(json);
	}
}
catch
{
	Console.WriteLine("[RO] No config file detected");
	Console.ReadLine();
	Environment.Exit(0);
}

// Select application
while(!appfound)
{
	Console.WriteLine("Please enter application to find (Excel or Word):");
	string read = Console.ReadLine();
	read = read.ToLower();

	switch(read)
	{
		case "word":
		client = new DiscordRpcClient(configuration.word_id);
		noun = "word document";
		appfound = true;
		input = read;
		break;
		case "excel":
		client = new DiscordRpcClient(configuration.excel_id);
		noun = "spreadsheet";
		appfound = true;
		input = read;
		break;
		default:
		Console.WriteLine("Invalid application!");
		break;
	}
}

// Initialise logging
client.Logger = new ConsoleLogger() { Level = LogLevel.Warning };

// Subscribe to events
client.OnReady += (sender, e) =>
{
	Console.WriteLine("[RPC] Received Ready from user {0}", e.User.Username);
};

client.OnPresenceUpdate += (sender, e) =>
{
	Console.WriteLine("[RPC] Received Update! {0}", e.Presence);
};

// Connect to the RPC
client.Initialize();

// Initial Activity
UpdatePresence();

// Main Loop
while(true)
{
	// Console.WriteLine($"Mem usage: {GC.GetTotalMemory(true):#,0} bytes");
}

void UpdatePresence()
{
	string title = "Idle";
	try
	{
		string windowTitle = Process.GetProcessesByName(input)[0].MainWindowTitle;
		Console.WriteLine("[RO] Found " + input + "! - " + windowTitle);
		title = windowTitle.Remove(windowTitle.Length - 8, 8);
	}
	catch
	{
		Console.WriteLine("[RO] No " + input + " Detected");
		Console.ReadLine();
		Environment.Exit(0);
	}

	client.SetPresence(new RichPresence()
	{
		Details = title,
		State = "DEVELOPMENT",
		Assets = new Assets()
		{
			LargeImageKey = input,
			LargeImageText = "Editing a " + noun + "..."//,
			//SmallImageKey = "image_small"
		}
	});
}

//Json
class Config
{
	public string word_id;
	public string excel_id;
}