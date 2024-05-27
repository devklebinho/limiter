from ctypes import cast, POINTER
from comtypes import CoInitialize, CoUninitialize, CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import psutil

# Inicializa o COM
CoInitialize()

# Obtém o dispositivo de áudio apenas uma vez para evitar repetidas chamadas
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume_interface = cast(interface, POINTER(IAudioEndpointVolume))

def get_system_volume():
    return volume_interface.GetMasterVolumeLevelScalar()

def set_system_volume(volume):
    volume_interface.SetMasterVolumeLevelScalar(volume, None)

def is_program_running(program_name):
    for process in psutil.process_iter(['pid', 'name']):
        if process.info['name'] == program_name:
            return True
    return False
    

def main():
    try:
        # Defina o volume desejado entre 0.0 e 1.0
        allowVolume = 0.01  # por exemplo, 40%
        set_system_volume(allowVolume)
        
        while True:
            current_volume = get_system_volume()
            if current_volume > 0.01:
                set_system_volume(allowVolume)
                #print("Volume ajustado para:", allowVolume)
            else:
                #print("Volume Ok")
                pass
    except KeyboardInterrupt:
        pass
    finally:
        # Libera o COM
        CoUninitialize()

if __name__ == "__main__":
    main()