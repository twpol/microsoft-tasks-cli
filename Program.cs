using Microsoft.Exchange.WebServices.Data;
using Microsoft.Extensions.Configuration;
using Task = Microsoft.Exchange.WebServices.Data.Task;
using TaskStatus = Microsoft.Exchange.WebServices.Data.TaskStatus;

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
    /// Command-line tool for manipulating Microsoft To Do/Tasks
    /// </summary>
    /// <param name="config">Path to configuration file</param>
    /// <param name="lists">Action: show all lists</param>
    /// <param name="tasks">Action: show all To Dos/Tasks in a list</param>
    /// <param name="createTask">Action: create a new To Do/Task</param>
    /// <param name="list">Specify the name or ID of a task list</param>
    /// <param name="key">Specify the substring to match an existing task</param>
    /// <param name="name">Specify the name of a task</param>
    /// <param name="body">Specify the body of a task</param>
    /// <param name="important">Specify that the task is important</param>
    /// <param name="complete">Specify that the task is completed</param>
    /// <param name="output">Specify the output format</param>
    static async System.Threading.Tasks.Task Main(FileInfo? config = null, bool lists = false, bool tasks = false, bool createTask = false, string list = "", string? key = null, string? name = null, string? body = null, bool? important = null, bool? complete = null, OutputFormat output = OutputFormat.Console)
    {
        config ??= new FileInfo("config.json");
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
            await CreateOrEditTask(LoadConfiguration(config), key, name, body, important, complete);
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
        var service = new ExchangeService(ExchangeVersion.Exchange2016)
        {
            Credentials = new WebCredentials(config["username"], config["password"])
        };
        service.AutodiscoverUrl(config["email"], redirectionUri => new Uri(redirectionUri).Scheme == "https");
        return service;
    }

    static async System.Threading.Tasks.Task Lists(IConfigurationRoot config)
    {
        var service = GetExchange(config);
        var taskFolder = await Retry("get tasks folder", () => Folder.Bind(service, WellKnownFolderName.Tasks));
        var lists = await Retry("get lists", () => taskFolder.FindFolders(new FolderView(1000)));
        if (Output == OutputFormat.Markdown)
        {
            Console.WriteLine($"# Lists");
            Console.WriteLine();
            Console.WriteLine($"- Tasks ({taskFolder.Id})");
        }
        foreach (var list in lists)
        {
            Console.WriteLine(FormatList(list));
        }
    }

    static async System.Threading.Tasks.Task Tasks(IConfigurationRoot config)
    {
        var service = GetExchange(config);
        var list = await GetList(service);
        var tasks = await Retry("get tasks", () => list.FindItems(new ItemView(1000) { PropertySet = PropertySet.IdOnly }));
        if (tasks.Any()) await service.LoadPropertiesForItems(tasks, new PropertySet(TaskSchema.Subject, TaskSchema.Body, TaskSchema.Importance, TaskSchema.IsComplete, TaskSchema.CompleteDate));
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

    static async System.Threading.Tasks.Task CreateOrEditTask(IConfigurationRoot config, string? key, string name, string? body, bool? important, bool? complete)
    {
        if (key != null && !name.Contains(key))
        {
            throw new InvalidDataException($"Task name does not contain key: {key}");
        }
        var service = GetExchange(config);
        var list = await GetList(service, always: true);
        var existingTasks = await Retry("get existing tasks", () =>
        {
            SearchFilter searchFilter = key == null
                ? new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(TaskSchema.Subject, name), new SearchFilter.IsEqualTo(TaskSchema.IsComplete, false))
                : new SearchFilter.ContainsSubstring(TaskSchema.Subject, key);
            return list.FindItems(searchFilter, new ItemView(1));
        });
        var task = existingTasks.FirstOrDefault() as Task;
        if (key == null && task != null)
        {
            Console.WriteLine($"WARNING: Duplicate task in {list.DisplayName}: {FormatTaskConsole(task)}");
            return;
        }
        var taskBody = new MessageBody(BodyType.Text, string.IsNullOrWhiteSpace(body) ? null : body);
        var taskImportance = important ?? false ? Importance.High : Importance.Normal;
        var taskStatus = complete ?? false ? TaskStatus.Completed : TaskStatus.NotStarted;
        if (task == null)
        {
            task = new Task(service)
            {
                Subject = name,
                Body = taskBody,
                Importance = taskImportance,
                Status = taskStatus,
            };
            await task.Save(list.Id);
            await task.Load();
            Console.WriteLine($"Created task in {list.DisplayName}: {FormatTaskConsole(task)}");
        }
        else
        {
            await task.Load(PropertySet.FirstClassProperties);
            if (task.Subject != name) task.Subject = name;
            if (task.Body.ToString() != taskBody.ToString()) task.Body = taskBody;
            if (important.HasValue && task.Importance != taskImportance) task.Importance = taskImportance;
            if (complete.HasValue && task.Status != taskStatus) task.Status = taskStatus;
            if (task.IsDirty)
            {
                await task.Update(ConflictResolutionMode.AlwaysOverwrite);
                await task.Load();
                Console.WriteLine($"Updated task in {list.DisplayName}: {FormatTaskConsole(task)}");
            }
        }
    }

    static async Task<Folder> GetList(ExchangeService service, bool always = false)
    {
        var taskFolder = await Retry("get tasks folder", () => Folder.Bind(service, WellKnownFolderName.Tasks));
        if (List == "") return taskFolder;
        try
        {
            return await Retry("get exact list", () => Folder.Bind(service, List));
        }
        catch (ServiceResponseException error) when (error.ErrorCode == ServiceError.ErrorInvalidId)
        {
            throw new InvalidDataException($"No list with ID: {List}");
        }
        catch (ServiceResponseException error) when (error.ErrorCode == ServiceError.ErrorInvalidIdMalformed)
        {
            // List is not an ID
        }
        var lists = await Retry("get list", () => taskFolder.FindFolders(new SearchFilter.ContainsSubstring(FolderSchema.DisplayName, List), new FolderView(1)));
        if (lists.TotalCount == 0 && always) return taskFolder;
        if (lists.TotalCount == 0) throw new InvalidDataException($"No list containing text: {List}");
        return lists.First();
    }

    static string FormatList(Folder list)
    {
        return Output switch
        {
            OutputFormat.Console => list.DisplayName,
            OutputFormat.Markdown => $"- {list.DisplayName} ({list.Id})",
            _ => throw new InvalidOperationException($"Unknown output format: {Output}"),
        };
    }

    static string FormatTask(Item item)
    {
        if (item is not Task task) return "";
        return Output switch
        {
            OutputFormat.Console => FormatTaskConsole(task),
            OutputFormat.Markdown => FormatTaskMarkdown(task),
            _ => throw new InvalidOperationException($"Unknown output format: {Output}"),
        };
    }

    static string FormatTaskConsole(Task task)
    {
        return $"[{(task.IsComplete ? "X" : " ")}] {(task.Importance == Importance.High ? "*" : " ")} {task.Subject}{(task.IsComplete ? $" (completion {task.CompleteDate:yyyy-MM-dd})" : "")}";
    }

    static string FormatTaskMarkdown(Task task)
    {
        return String.Join("\n  ", Split("\n", task.Body.ToString()).Prepend($"- [{(task.IsComplete ? "X" : " ")}] {task.Subject}{(task.Importance == Importance.High ? " [important:: true]" : "")}{(task.IsComplete ? $" [completion:: {task.CompleteDate:yyyy-MM-dd}]" : "")}"));
    }

    static string[] Split(string separator, string text)
    {
        if (String.IsNullOrWhiteSpace(text)) return Array.Empty<string>();
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
