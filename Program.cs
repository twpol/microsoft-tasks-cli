using Microsoft.Exchange.WebServices.Data;
using Microsoft.Extensions.Configuration;
using Task = System.Threading.Tasks.Task;

class Program
{
    /// <summary>
    /// A command-line tool for manipulating Microsoft To Do/Tasks
    /// </summary>
    /// <param name="config">Path to configuration file</param>
    /// <param name="lists">Action: show all lists</param>
    /// <param name="tasks">Action: show all To Dos/Tasks in a list</param>
    /// <param name="createTask">Action: create a new To Do/Task</param>
    /// <param name="list">Specify the name or ID of a task list</param>
    /// <param name="title">Specify the title of a task</param>
    /// <param name="description">Specify the description of a task</param>
    static async Task Main(FileInfo? config = null, bool lists = false, bool tasks = false, bool createTask = false, string? list = null, string? title = null, string? description = null)
    {
        if (config == null) config = new FileInfo("config.json");
        if (lists)
        {
            await Lists(LoadConfiguration(config));
        }
        else if (tasks)
        {
            ArgumentNullException.ThrowIfNull(list);
            await Tasks(LoadConfiguration(config), list);
        }
        else if (createTask)
        {
            ArgumentNullException.ThrowIfNull(list);
            ArgumentNullException.ThrowIfNull(title);
            CreateTask(LoadConfiguration(config), list, title, description ?? "");
        }
        else
        {
            throw new InvalidOperationException($"No valid action specified");
        }
    }

    static IConfigurationRoot LoadConfiguration(FileInfo config)
    {
        return new ConfigurationBuilder()
            .AddJsonFile(config.FullName, true)
            .Build();
    }

    static ExchangeService GetExchange(IConfigurationRoot config)
    {
        var service = new ExchangeService(ExchangeVersion.Exchange2016);
        service.Credentials = new WebCredentials(config["username"], config["password"]);
        service.AutodiscoverUrl(config["email"], redirectionUri => new Uri(redirectionUri).Scheme == "https");
        return service;
    }

    static async Task Lists(IConfigurationRoot config)
    {
        var service = GetExchange(config);
        var taskFolder = await Retry("get tasks folder", () => Folder.Bind(service, WellKnownFolderName.Tasks));
        var lists = await Retry("get lists", () => taskFolder.FindFolders(new FolderView(1000)));
        foreach (var list in lists)
        {
            Console.WriteLine(list.DisplayName);
        }
    }

    static async Task Tasks(IConfigurationRoot config, string listName)
    {
        var service = GetExchange(config);
        var taskFolder = await Retry("get tasks folder", () => Folder.Bind(service, WellKnownFolderName.Tasks));
        var lists = await Retry("get list", () => taskFolder.FindFolders(new SearchFilter.ContainsSubstring(FolderSchema.DisplayName, listName), new FolderView(1)));
        if (lists.TotalCount == 0) throw new InvalidDataException($"No list containing text: {listName}");
        var tasks = await Retry("get tasks", () => lists.First().FindItems(new ItemView(1000)));
        foreach (var task in tasks)
        {
            Console.WriteLine(task.Subject);
        }
    }

    static void CreateTask(IConfigurationRoot config, string list, string title, string description)
    {
        var service = GetExchange(config);
        // TODO: Find list
        // TODO: Find task
        // TODO: Create task
    }

    static async Task<T> Retry<T>(string name, Func<Task<T>> action)
    {
        while (true)
        {
            try
            {
                return await action();
            }
            catch (ServerBusyException error)
            {
                Console.WriteLine($"Retry of {name} due to server busy (back off for {error.BackOffMilliseconds} ms)");
                Thread.Sleep(error.BackOffMilliseconds);
            }
        }
    }
}
