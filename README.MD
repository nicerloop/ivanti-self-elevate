# Launch self-elevated commands using Ivanti Application Control

This is a simple command to launch elevated processes using Ivanti Application Control shell integration.

## Ivanti Application Control Self-Elevation Shell Extension

The core is directly taken from [ykoster/Start-ProcessAMSelfElevate.psm1](https://gist.github.com/ykoster/635b3639bdd5f6c1d9c5283ae8460ed6).

The original script
> uses the AMShellIntegration.AMShellContextMenu COM component to invoke the Self-Elevation functionality of Ivanti AppSense Application 
Manager.
>
> When enabled, Self-Elevation can be used to run (selected) applications with Administrator privileges. Generally, this functionality is only exposed from Windows Explorer. Using this module, it is possible to call this component directly (with arbitrary arguments).

The C# class is reused as-is, a simplistic `Main` method is added to launch the program.

## Usage

It is meant to be compatible with [gsudo](https://github.com/gerardog/gsudo) as used by [WingetUI](https://github.com/marticliment/WingetUI).

### Version

```
$> gsudo -v
ivanti-gsudo 0.0.1
```

### Cache

Cache is not implemented, only compatible arguments parsing is.

```
$> gsudo cache ...
Info: No cache inplemented.
```

```
$> gsudo -k
Info: No cache inplemented.
```

### Run elevated command

```
$> gsudo notepad C:\Windows\System32\drivers\etc\hosts
```
A dialog for self-elevation reason auditing is displayed, and the process is launched upon dialog validation.
