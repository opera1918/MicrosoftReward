from win32com import client
import subprocess

def kill_process(program):
    """
    Get the process ID of Microsoft Edge
    """
    for task in client.GetObject('winmgmts:').InstancesOf('win32_process'):
        if task.name == f'{program}.exe':
            subprocess.check_output(f"Taskkill /PID {task.processID} /F")
            print(f"[KILLED] {task.name}: {task.processID} terminated")
            break
if __name__ == '__main__':
    kill_process('Code')
    pass

