import sys
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import subprocess
import os
from rich.console import Console
from rich.logging import RichHandler
from rich.panel import Panel
from rich.live import Live
from rich.table import Table
from rich.spinner import Spinner
from rich.traceback import install
import logging
from datetime import datetime
import psutil
import signal
import traceback

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

# Signal names mapping
SIGNAL_NAMES = {
    -1: "SIGHUP",     # Hangup
    -2: "SIGINT",     # Interrupt
    -3: "SIGQUIT",    # Quit
    -4: "SIGILL",     # Illegal instruction
    -6: "SIGABRT",    # Abort
    -8: "SIGFPE",     # Floating point exception
    -9: "SIGKILL",    # Kill
    -11: "SIGSEGV",   # Segmentation fault
    -13: "SIGPIPE",   # Broken pipe
    -15: "SIGTERM",   # Termination
}

def get_signal_description(return_code):
    """Get a human-readable description of a signal-based return code."""
    if return_code >= 0:
        return f"Process exited with code {return_code}"
        
    signal_num = abs(return_code)
    signal_name = SIGNAL_NAMES.get(return_code, f"Unknown signal ({signal_num})")
    
    descriptions = {
        "SIGSEGV": "Memory access violation (segmentation fault). This might be caused by:\n"
                   "• Accessing invalid memory addresses\n"
                   "• Stack overflow\n"
                   "• Memory corruption",
        "SIGABRT": "Process abort signal. This might be caused by:\n"
                   "• Unhandled exceptions\n"
                   "• Critical resource failures\n"
                   "• Internal assertion failures",
        "SIGFPE": "Floating point exception. This might be caused by:\n"
                  "• Division by zero\n"
                  "• Invalid floating point operation",
        "SIGILL": "Illegal instruction. This might be caused by:\n"
                  "• Corrupted executable memory\n"
                  "• Invalid code generation",
        "SIGPIPE": "Broken pipe. This might be caused by:\n"
                   "• Writing to a closed pipe/socket\n"
                   "• Network connection issues",
    }
    
    description = descriptions.get(signal_name, "Unknown error occurred")
    return f"{signal_name}: {description}"

def get_process_info(process):
    """Get detailed information about a process."""
    try:
        if process and process.poll() is None:  # If process is still running
            proc = psutil.Process(process.pid)
            memory_info = proc.memory_info()
            cpu_percent = proc.cpu_percent(interval=0.1)
            return {
                'memory_rss': memory_info.rss / (1024 * 1024),  # Convert to MB
                'memory_vms': memory_info.vms / (1024 * 1024),  # Convert to MB
                'cpu_percent': cpu_percent,
                'threads': proc.num_threads(),
                'status': proc.status()
            }
    except (psutil.NoSuchProcess, psutil.AccessDenied, Exception) as e:
        logging.warning(f"Failed to get process info: {e}")
    return None

class RestartHandler(FileSystemEventHandler):
    def __init__(self):
        self.process = None
        self.start_time = None
        self.restart_count = 0
        self.last_crash_time = None
        self.crash_count = 0
        self.crash_history = []
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

            # Create log directory if it doesn't exist
            log_dir = os.path.join(os.path.expanduser("~"), ".edge_tts_gui", "logs")
            os.makedirs(log_dir, exist_ok=True)
            
            # Create log file for this session
            log_file = os.path.join(log_dir, f"edge_tts_gui_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
            
            # Redirect stdout and stderr to both console and file
            self.process = subprocess.Popen(
                [sys.executable, 'main.py'],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                bufsize=1,
                universal_newlines=True
            )

            # Start threads to monitor output and process
            self._monitor_output(self.process)
            self._monitor_process(self.process)

        except Exception as e:
            console.print(f"[red]Error:[/red] Failed to start application: {str(e)}", style="bold red")
            console.print(Panel(traceback.format_exc(), title="[red]Traceback[/red]", border_style="red"))

    def _monitor_process(self, process):
        """Monitor process resources and health."""
        def monitor():
            while process.poll() is None:  # While process is running
                try:
                    info = get_process_info(process)
                    if info:
                        # Log resource usage
                        logging.debug(
                            f"Process stats - Memory: {info['memory_rss']:.1f}MB, "
                            f"CPU: {info['cpu_percent']:.1f}%, "
                            f"Threads: {info['threads']}"
                        )
                        
                        # Check for potential memory leaks
                        if info['memory_rss'] > 1000:  # Warning if using more than 1GB RAM
                            logging.warning(f"High memory usage detected: {info['memory_rss']:.1f}MB")
                            
                    time.sleep(5)  # Check every 5 seconds
                except Exception as e:
                    logging.error(f"Error monitoring process: {e}")
                    break

        import threading
        threading.Thread(target=monitor, daemon=True).start()

    def _monitor_output(self, process):
        def read_output(pipe, is_error):
            try:
                for line in pipe:
                    line = line.strip()
                    if line:
                        if is_error:
                            console.print(f"[red]Error:[/red] {line}", style="red")
                            # Log errors to file
                            logging.error(line)
                        else:
                            console.print(f"[dim]{datetime.now().strftime('%H:%M:%S')}[/dim] {line}")
                            # Log standard output to file
                            logging.info(line)
            except Exception as e:
                console.print(f"[red]Error:[/red] Output monitoring failed: {str(e)}", style="bold red")
                logging.error(f"Output monitoring failed: {e}")

        import threading
        threading.Thread(target=read_output, args=(process.stdout, False), daemon=True).start()
        threading.Thread(target=read_output, args=(process.stderr, True), daemon=True).start()

    def handle_crash(self, return_code):
        """Handle application crash with detailed reporting."""
        now = datetime.now()
        
        # Update crash statistics
        self.crash_count += 1
        self.crash_history.append({
            'time': now,
            'return_code': return_code,
            'uptime': (now - self.start_time).total_seconds() if self.start_time else 0
        })
        
        # Get crash details
        crash_description = get_signal_description(return_code)
        
        # Create detailed crash report
        crash_info = [
            f"[red]Application Crash Report[/red]",
            f"Time: {now.strftime('%Y-%m-%d %H:%M:%S')}",
            f"Return Code: {return_code}",
            f"Description: {crash_description}",
            f"Uptime: {(now - self.start_time).total_seconds():.1f} seconds",
            f"Total Crashes: {self.crash_count}",
            "",
            "[yellow]Recent Process Information:[/yellow]"
        ]
        
        # Add process info if available
        proc_info = get_process_info(self.process)
        if proc_info:
            crash_info.extend([
                f"Memory Usage: {proc_info['memory_rss']:.1f}MB",
                f"CPU Usage: {proc_info['cpu_percent']:.1f}%",
                f"Thread Count: {proc_info['threads']}"
            ])
        
        # Display crash report
        console.print(Panel(
            "\n".join(crash_info),
            title="Crash Report",
            border_style="red"
        ))
        
        # Log crash information
        logging.error(f"Application crash detected: {crash_description}")
        
        # Check for crash frequency
        if len(self.crash_history) >= 3:
            recent_crashes = self.crash_history[-3:]
            time_span = (recent_crashes[-1]['time'] - recent_crashes[0]['time']).total_seconds()
            if time_span < 60:  # 3 crashes within 1 minute
                console.print(Panel(
                    "[red]Warning: High Crash Frequency Detected[/red]\n"
                    "The application has crashed multiple times in a short period.\n"
                    "This might indicate a serious issue that needs attention.",
                    border_style="red"
                ))
                logging.critical("High crash frequency detected!")

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
                        # Handle crash with detailed reporting
                        event_handler.handle_crash(return_code)
                        # Attempt to restart
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
        console.print(Panel(traceback.format_exc(), title="[red]Traceback[/red]", border_style="red"))
    finally:
        observer.join()
        console.print(Panel.fit(
            "[bold green]Development server stopped[/bold green]",
            border_style="green"
        )) 