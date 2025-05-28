import sys
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import subprocess
import os
import psutil
from rich.console import Console
from rich.logging import RichHandler
from rich.panel import Panel
from rich.live import Live
from rich.table import Table
from rich.spinner import Spinner
from rich.traceback import install
import logging
from datetime import datetime
import signal

# Install rich traceback handler
install(show_locals=True)

# Initialize Rich console
console = Console()

# Set up logging with Rich
logging.basicConfig(
    level=logging.INFO,
    format="%(message)s",
    handlers=[RichHandler(rich_tracebacks=True, markup=True)]
)

# Signal names dictionary
SIGNAL_NAMES = {
    -11: "SIGSEGV (Segmentation Fault)",
    -6: "SIGABRT (Abort)",
    -4: "SIGILL (Illegal Instruction)",
    -5: "SIGTRAP (Trace/Breakpoint Trap)",
    -8: "SIGFPE (Floating Point Exception)",
    -7: "SIGBUS (Bus Error)",
}

class RestartHandler(FileSystemEventHandler):
    def __init__(self):
        self.process = None
        self.start_time = None
        self.restart_count = 0
        console.print(Panel.fit(
            "[bold blue]EdgeTTS-GUI Development Server[/bold blue]\n"
            "[dim]Auto-reloading enabled[/dim]",
            border_style="blue"
        ))
        self.start_application()

    def start_application(self):
        try:
            if self.process:
                console.print("[yellow]→[/yellow] Terminating previous instance...", style="bold")
                # Get process info before terminating
                try:
                    proc = psutil.Process(self.process.pid)
                    memory_info = proc.memory_info()
                    cpu_percent = proc.cpu_percent()
                    console.print(f"[dim]Process stats before termination:[/dim]")
                    console.print(f"[dim]- Memory usage: {memory_info.rss / 1024 / 1024:.2f} MB[/dim]")
                    console.print(f"[dim]- CPU usage: {cpu_percent}%[/dim]")
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass

                self.process.terminate()
                try:
                    self.process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    console.print("[red]![/red] Process did not terminate gracefully, forcing kill...", style="bold")
                    self.process.kill()

            self.restart_count += 1
            self.start_time = datetime.now()
            
            # Create a status table
            table = Table(show_header=False, box=None)
            table.add_row(
                "[green]●[/green]", 
                f"Starting application (Restart #{self.restart_count})"
            )
            console.print(table)

            # Redirect stdout and stderr to pipe to capture output
            self.process = subprocess.Popen(
                [sys.executable, 'main.py'],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                bufsize=1,
                universal_newlines=True
            )

            # Start threads to monitor output
            self._monitor_output(self.process)

        except Exception as e:
            console.print(f"[red]Error:[/red] Failed to start application: {str(e)}", style="bold red")

    def _monitor_output(self, process):
        def read_output(pipe, is_error):
            try:
                for line in pipe:
                    line = line.strip()
                    if line:
                        if is_error:
                            console.print(f"[red]Error:[/red] {line}", style="red")
                        else:
                            console.print(f"[dim]{datetime.now().strftime('%H:%M:%S')}[/dim] {line}")
            except Exception as e:
                console.print(f"[red]Error:[/red] Output monitoring failed: {str(e)}", style="bold red")

        import threading
        threading.Thread(target=read_output, args=(process.stdout, False), daemon=True).start()
        threading.Thread(target=read_output, args=(process.stderr, True), daemon=True).start()

    def on_modified(self, event):
        if event.src_path.endswith('.py'):
            console.print(
                Panel(
                    f"[yellow]File changed:[/yellow] {os.path.basename(event.src_path)}",
                    border_style="yellow",
                    expand=False
                )
            )
            self.start_application()

if __name__ == "__main__":
    path = '.'
    event_handler = RestartHandler()
    observer = Observer()

    try:
        observer.schedule(event_handler, path, recursive=False)
        observer.start()
        
        with Live(
            Panel(
                "[green]Development server is running[/green]\n[dim]Press Ctrl+C to stop[/dim]",
                border_style="green"
            ),
            refresh_per_second=4,
            transient=True
        ) as live:
            while True:
                time.sleep(1)
                # Check if the process has crashed
                if event_handler.process and event_handler.process.poll() is not None:
                    return_code = event_handler.process.poll()
                    if return_code != 0:
                        # Get crash duration
                        crash_duration = datetime.now() - event_handler.start_time
                        
                        # Create detailed crash report
                        crash_info = [
                            f"[red]Application crashed after running for {crash_duration.total_seconds():.1f} seconds[/red]",
                            f"[yellow]Return code: {return_code}[/yellow]"
                        ]
                        
                        # Add signal information if it's a signal-related crash
                        if return_code < 0:
                            signal_name = SIGNAL_NAMES.get(return_code, "Unknown signal")
                            crash_info.append(f"[yellow]Signal: {signal_name}[/yellow]")
                            
                            if return_code == -11:  # SIGSEGV
                                crash_info.append("[dim]A segmentation fault typically indicates:")
                                crash_info.append("- Accessing invalid memory addresses")
                                crash_info.append("- Stack overflow")
                                crash_info.append("- Null pointer dereference[/dim]")
                        
                        # Create and display the crash panel
                        crash_panel = Panel(
                            "\n".join(crash_info),
                            title="[red]Crash Report[/red]",
                            border_style="red"
                        )
                        console.print(crash_panel)
                        
                        console.print("[yellow]Attempting to restart...[/yellow]")
                        event_handler.start_application()

    except KeyboardInterrupt:
        console.print("\n[yellow]Shutting down...[/yellow]", style="bold")
        if event_handler.process:
            event_handler.process.terminate()
            try:
                event_handler.process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                event_handler.process.kill()
        observer.stop()
    except Exception as e:
        console.print(f"[red]Fatal Error:[/red] {str(e)}", style="bold red")
    finally:
        observer.join()
        console.print(Panel.fit(
            "[bold green]Development server stopped[/bold green]",
            border_style="green"
        )) 