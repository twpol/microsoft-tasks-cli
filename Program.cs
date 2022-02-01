using Microsoft.Exchange.WebServices.Data;
using Microsoft.Extensions.Configuration;
using Task = Microsoft.Exchange.WebServices.Data.Task;

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
    /// <param name="name">Specify the name of a task</param>
    /// <param name="body">Specify the body of a task</param>
    /// <param name="important">Specify that the task is important</param>
    static async System.Threading.Tasks.Task Main(FileInfo? config = null, bool lists = false, bool tasks = false, bool createTask = false, string? list = null, string? name = null, string? body = null, bool important = false)
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
            ArgumentNullException.ThrowIfNull(name);
            await CreateTask(LoadConfiguration(config), list, name, body ?? "", important);
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

    static async System.Threading.Tasks.Task Lists(IConfigurationRoot config)
    {
        var service = GetExchange(config);
        var taskFolder = await Retry("get tasks folder", () => Folder.Bind(service, WellKnownFolderName.Tasks));
        var lists = await Retry("get lists", () => taskFolder.FindFolders(new FolderView(1000)));
        foreach (var list in lists)
        {
            Console.WriteLine(list.DisplayName);
        }
    }

    static async System.Threading.Tasks.Task Tasks(IConfigurationRoot config, string listName)
    {
        var service = GetExchange(config);
        var list = await GetList(service, listName);
        var tasks = await Retry("get tasks", () => list.FindItems(new ItemView(1000)));
        foreach (var task in tasks)
        {
            Console.WriteLine(FormatTask(task));
        }
    }

    static async System.Threading.Tasks.Task CreateTask(IConfigurationRoot config, string listName, string name, string body, bool important)
    {
        var service = GetExchange(config);
        var list = await GetList(service, listName, always: true);
        var existingTasks = await Retry("get existing tasks", () => list.FindItems(new SearchFilter.SearchFilterCollection(
            LogicalOperator.And,
            new SearchFilter.IsEqualTo(TaskSchema.Subject, name),
            new SearchFilter.IsEqualTo(TaskSchema.IsComplete, false)
        ), new ItemView(1)));
        if (existingTasks.TotalCount > 0)
        {
            Console.WriteLine($"WARNING: Duplicate task in {list.DisplayName}: {FormatTask(existingTasks.First())}");
            return;
        }
        var task = new Task(service);
        task.Subject = name;
        task.Body = body;
        task.Importance = important ? Importance.High : Importance.Normal;
        await task.Save(list.Id);
        await task.Load();
        Console.WriteLine($"Created task in {list.DisplayName}: {FormatTask(task)}");
    }

    static async Task<Folder> GetList(ExchangeService service, string listName, bool always = false)
    {
        var taskFolder = await Retry("get tasks folder", () => Folder.Bind(service, WellKnownFolderName.Tasks));
        if (listName == "") return taskFolder;
        var lists = await Retry("get list", () => taskFolder.FindFolders(new SearchFilter.ContainsSubstring(FolderSchema.DisplayName, listName), new FolderView(1)));
        if (lists.TotalCount == 0 && always) return taskFolder;
        if (lists.TotalCount == 0) throw new InvalidDataException($"No list containing text: {listName}");
        return lists.First();
    }

    static string FormatTask(Item item)
    {
        var task = item as Task;
        return $"[{(task?.IsComplete ?? false ? "X" : " ")}] {(item.Importance == Importance.High ? "*" : " ")} {item.Subject}";
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
