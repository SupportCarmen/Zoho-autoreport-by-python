from rich.console import Console
from datetime import datetime

try:
    from zoneinfo import ZoneInfo
except ImportError:
    pass

class BotLogger:
    def __init__(self):
        self.console = Console()
        
    def _time(self):
        try:
            return datetime.now(ZoneInfo('Asia/Bangkok')).strftime('%H:%M:%S')
        except:
            return datetime.now().strftime('%H:%M:%S')

    def info(self, msg):
        self.console.print(f"[[cyan]{self._time()}[/cyan]] [cyan]{msg}[/cyan]")

    def success(self, msg):
        self.console.print(f"[[cyan]{self._time()}[/cyan]] [bold green]{msg}[/bold green]")

    def warning(self, msg):
        self.console.print(f"[[cyan]{self._time()}[/cyan]] [bold yellow]{msg}[/bold yellow]")

    def error(self, msg):
        self.console.print(f"[[cyan]{self._time()}[/cyan]] [bold red]{msg}[/bold red]")

    def step(self, msg):
        self.console.print()
        self.console.print(f"[[cyan]{self._time()}[/cyan]] [bold magenta]🚀 {msg}[/bold magenta]")
        
    def print(self, msg=""):
        self.console.print(msg)

log = BotLogger()
