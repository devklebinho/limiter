from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

def set_system_volume(volume):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume_interface = cast(interface, POINTER(IAudioEndpointVolume))
    volume_interface.SetMasterVolumeLevelScalar(volume, None)

def main():
    # Defina o volume desejado entre 0.0 e 1.0
    volume = 0.0  # por exemplo, 0.5 representa 50%
    set_system_volume(volume)

if __name__ == "__main__":
    main()