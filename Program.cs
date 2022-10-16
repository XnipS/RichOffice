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
void SelectApp ()
{
	while(!appfound)
	{
		Console.WriteLine("> Please enter application to find (Excel or Word):");
		string read = Console.ReadLine();
		read = read.ToLower();

		switch(read)
		{
			case "word":
			client = new DiscordRpcClient(configuration.word_id);
			noun = "word document";
			StartClient(read);
			break;
			case "excel":
			client = new DiscordRpcClient(configuration.excel_id);
			noun = "spreadsheet";
			StartClient(read);
			break;
			default:
			Console.WriteLine("> Invalid application!");
			break;
		}
	}
}

void StartClient (string read)
{
	appfound = true;
	input = read;

	// Initialise logging
	client.Logger = new ConsoleLogger() { Level = LogLevel.Warning };

	// Subscribe to events
	client.OnReady += (sender, e) =>
	{
		Console.WriteLine("\r[RPC] Received Ready from user {0}", e.User.Username);
		Console.Write("[RO] Ready! > ");
	};

	client.OnPresenceUpdate += (sender, e) =>
	{
		Console.WriteLine("\r[RPC] Received Update! {0}", e.Presence);
		Console.Write("[RO] Ready! > ");
	};

	// Connect to the RPC
	client.Initialize();

}

// Command Loop
while(true)
{
	Console.Write("[RO] Ready! > ");
	string command = Console.ReadLine();
	command = command.ToLower();
	switch(command)
	{
		case "memory" or "mem":
		Console.WriteLine("> Mem usage: {GC.GetTotalMemory(true):#,0} bytes");
		break;
		case "update":
		UpdatePresence();
		break;
		case "swap" or "select":
		appfound = false;
		SelectApp();
		break;
		case "exit":
		Environment.Exit(0);
		break;
		default:
		Console.WriteLine("> Invalid command!");
		break;
	}
}

// Presence Update
void UpdatePresence()
{
	string title = "Idle";
	try
	{
		string windowTitle;
		if(input == "word")
		{
			windowTitle = Process.GetProcessesByName("win" + input)[0].MainWindowTitle;
		}
		else
		{
			windowTitle = Process.GetProcessesByName(input)[0].MainWindowTitle;
		}
		Console.WriteLine("[RO] Found " + input + "! - " + windowTitle);
		title = windowTitle.Remove(windowTitle.Length - 8, 8);
	}
	catch
	{
		if(input != " ")
		{
			Console.WriteLine("[RO] No application selected...");
		}
		else
		{
			Console.WriteLine("[RO] No " + input + " detected...");
		}
		return;
	}

	client.SetPresence(new RichPresence()
	{
		Details = title,
		State = "Editing...",
		Timestamps = new Timestamps()
		{
			Start = DateTime.UtcNow,
			End = null
		},
		Assets = new Assets()
		{
			LargeImageKey = input,
			LargeImageText = "Editing a " + noun + "..."
		}
	});
}

//Json
class Config
{
	public string word_id;
	public string excel_id;
}