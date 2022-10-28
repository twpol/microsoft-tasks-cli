# Microsoft Tasks CLI

Command-line tool for manipulating Microsoft To Do/Tasks.

## Usage

```
dotnet run -- [options]
```
```
MicrosoftTasksCLI [options]
```

## Options

- `--config <config>`

  Path to configuration file [] (required).

- `--lists`

  Action: show all lists.

- `--tasks`

  Action: show all To Dos/Tasks in a list.

- `--create-task`

  Action: create a new To Do/Task.

- `--list <list>`

  Specify the name or ID of a task list [] (required).

- `--name <name>`

  Specify the name of a task [] (required).

- `--body <body>`

  Specify the body of a task [] (required).

- `--important`

  Specify that the task is important.

- `--output <Console|Markdown>`

  Specify the output format (default: Console).
