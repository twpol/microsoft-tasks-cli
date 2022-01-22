using Microsoft.Extensions.Configuration;

class Program
{
    /// <summary>
    /// A command-line tool for manipulating Microsoft To Do/Tasks
    /// </summary>
    /// <param name="config">Path to configuration file</param>
    /// <param name="createTask">Action: create a new To Do/Task</param>
    /// <param name="list">Specify the name or ID of a task list</param>
    /// <param name="title">Specify the title of a task</param>
    /// <param name="description">Specify the description of a task</param>
    static void Main(FileInfo? config = null, bool createTask = false, string? list = null, string? title = null, string? description = null)
    {
        if (config == null) config = new FileInfo("config.json");
        if (createTask)
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

    static void CreateTask(IConfigurationRoot config, string list, string title, string description)
    {
        Console.WriteLine($"Create Task");
        Console.WriteLine($"List:        {list}");
        Console.WriteLine($"Title:       {title}");
        Console.WriteLine($"Description: {description}");
    }
}
