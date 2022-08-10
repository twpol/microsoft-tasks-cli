using Microsoft.Exchange.WebServices.Data;
using Microsoft.Extensions.Configuration;
using Task = Microsoft.Exchange.WebServices.Data.Task;

class Program
{
    enum OutputFormat
    {
        Console,
        Markdown,
    }

    static string List = "";
    static OutputFormat Output = OutputFormat.Console;

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
    /// <param name="output">Specify the output format</param>
    static async System.Threading.Tasks.Task Main(FileInfo? config = null, bool lists = false, bool tasks = false, bool createTask = false, string list = "", string? name = null, string? body = null, bool important = false, OutputFormat output = OutputFormat.Console)
    {
        if (config == null) config = new FileInfo("config.json");
        List = list;
        Output = output;
        if (lists)
        {
            await Lists(LoadConfiguration(config));
        }
        else if (tasks)
        {
            await Tasks(LoadConfiguration(config));
        }
        else if (createTask)
        {
            ArgumentNullException.ThrowIfNull(name);
            await CreateTask(LoadConfiguration(config), name, body ?? "", important);
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

    static async System.Threading.Tasks.Task Tasks(IConfigurationRoot config)
    {
        var service = GetExchange(config);
        var list = await GetList(service);
        var tasks = await Retry("get tasks", () => list.FindItems(new ItemView(1000) { PropertySet = PropertySet.IdOnly }));
        if (tasks.Any()) await service.LoadPropertiesForItems(tasks, new PropertySet(TaskSchema.Subject, TaskSchema.Body, TaskSchema.Importance, TaskSchema.IsComplete));
        if (Output == OutputFormat.Markdown)
        {
            Console.WriteLine($"# {list.DisplayName}");
            Console.WriteLine();
        }
        foreach (var task in tasks)
        {
            Console.WriteLine(FormatTask(task));
        }
    }

    static async System.Threading.Tasks.Task CreateTask(IConfigurationRoot config, string name, string body, bool important)
    {
        var service = GetExchange(config);
        var list = await GetList(service, always: true);
        var existingTasks = await Retry("get existing tasks", () => list.FindItems(new SearchFilter.SearchFilterCollection(
            LogicalOperator.And,
            new SearchFilter.IsEqualTo(TaskSchema.Subject, name),
            new SearchFilter.IsEqualTo(TaskSchema.IsComplete, false)
        ), new ItemView(1)));
        if (existingTasks.TotalCount > 0)
        {
            Console.WriteLine($"WARNING: Duplicate task in {list.DisplayName}: {FormatTaskConsole(existingTasks.First() as Task)}");
            return;
        }
        var task = new Task(service);
        task.Subject = name;
        task.Body = body;
        task.Importance = important ? Importance.High : Importance.Normal;
        await task.Save(list.Id);
        await task.Load();
        Console.WriteLine($"Created task in {list.DisplayName}: {FormatTaskConsole(task)}");
    }

    static async Task<Folder> GetList(ExchangeService service, bool always = false)
    {
        var taskFolder = await Retry("get tasks folder", () => Folder.Bind(service, WellKnownFolderName.Tasks));
        if (List == "") return taskFolder;
        var list = await Retry("get exact list", () => Folder.Bind(service, List));
        if (list != null) return list;
        var lists = await Retry("get list", () => taskFolder.FindFolders(new SearchFilter.ContainsSubstring(FolderSchema.DisplayName, List), new FolderView(1)));
        if (lists.TotalCount == 0 && always) return taskFolder;
        if (lists.TotalCount == 0) throw new InvalidDataException($"No list containing text: {List}");
        return lists.First();
    }

    static string FormatTask(Item item)
    {
        var task = item as Task;
        if (task == null) return "";
        switch (Output)
        {
            case OutputFormat.Console:
                return FormatTaskConsole(task);
            case OutputFormat.Markdown:
                return FormatTaskMarkdown(task);
            default:
                throw new InvalidOperationException($"Unknown output format: {Output}");
        }
    }

    static string FormatTaskConsole(Task? task)
    {
        return $"[{(task?.IsComplete ?? false ? "X" : " ")}] {(task?.Importance == Importance.High ? "*" : " ")} {task?.Subject}";
    }

    static string FormatTaskMarkdown(Task task)
    {
        return String.Join("\n  ", Split("\n", task.Body.ToString()).Prepend($"- [{(task.IsComplete ? "X" : " ")}] {task.Subject}{(task.Importance == Importance.High ? " [important:: true]" : "")}"));
    }

    static string[] Split(string separator, string text)
    {
        if (String.IsNullOrWhiteSpace(text)) return new string[0];
        return text.Split(separator);
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
