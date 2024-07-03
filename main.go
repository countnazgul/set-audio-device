package main

import (
	"fmt"
	"os"
	"syscall"
	"unsafe"

	"github.com/gen2brain/beeep"
	"github.com/go-ole/go-ole"
	"github.com/moutend/go-wca/pkg/wca"
)

func main() {
	err := ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED)
	if err != nil {
		panic(err)
	}

	args := os.Args
	if len(args) == 1 {
		listDevices()
		defer ole.CoUninitialize()
		os.Exit(0)
	}

	if len(args) == 2 {
		if args[1] == "current" {
			printCurrentDevice()
		} else {
			setAudioDeviceByID(args[1])
		}

		defer ole.CoUninitialize()
		os.Exit(0)
	}

	fmt.Printf("Too many arguments! Please provide only one")
	defer ole.CoUninitialize()

	os.Exit(1)
}

func listDevices() {
	devices := getAllDevices()

	for k, v := range devices {
		fmt.Printf("%s\n\t%s\n\n", k, v)
	}
}

func getDeviceName(deviceID string) string {
	devices := getAllDevices()
	deviceName := ""

	for k, v := range devices {
		if k == deviceID {
			deviceName = v
		}
	}

	return deviceName
}

func getAllDevices() map[string]string {

	var mmde *wca.IMMDeviceEnumerator

	if err := wca.CoCreateInstance(
		wca.CLSID_MMDeviceEnumerator,
		0,
		wca.CLSCTX_ALL,
		wca.IID_IMMDeviceEnumerator,
		&mmde,
	); err != nil {
		panic(err)
	}

	defer mmde.Release()

	var devicesCollection *wca.IMMDeviceCollection
	if err := mmde.EnumAudioEndpoints(wca.ERender, wca.DEVICE_STATE_ACTIVE, &devicesCollection); err != nil {
		panic(err)
	}

	var devicesCount uint32
	devicesCollection.GetCount(&devicesCount)

	defer devicesCollection.Release()

	myDict := make(map[string]string)

	for i := uint32(0); i < (devicesCount); i++ {
		var mmd *wca.IMMDevice
		if err := devicesCollection.Item(i, &mmd); err != nil {
			panic(err)
		}

		var deviceId string
		mmd.GetId(&deviceId)

		defer mmd.Release()

		var ps *wca.IPropertyStore
		if err := mmd.OpenPropertyStore(wca.STGM_READ, &ps); err != nil {
			panic(err)
		}
		defer ps.Release()

		var pv wca.PROPVARIANT
		if err := ps.GetValue(&wca.PKEY_Device_FriendlyName, &pv); err != nil {
			panic(err)
		}

		myDict[deviceId] = pv.String()
	}

	return myDict
}

type IPolicyConfigVista struct {
	ole.IUnknown
}

type IPolicyConfigVistaVtbl struct {
	ole.IUnknownVtbl
	GetMixFormat          uintptr
	GetDeviceFormat       uintptr
	SetDeviceFormat       uintptr
	GetProcessingPeriod   uintptr
	SetProcessingPeriod   uintptr
	GetShareMode          uintptr
	SetShareMode          uintptr
	GetPropertyValue      uintptr
	SetPropertyValue      uintptr
	SetDefaultEndpoint    uintptr
	SetEndpointVisibility uintptr
}

func (v *IPolicyConfigVista) VTable() *IPolicyConfigVistaVtbl {
	return (*IPolicyConfigVistaVtbl)(unsafe.Pointer(v.RawVTable))
}

func (v *IPolicyConfigVista) SetDefaultEndpoint(deviceID string, eRole wca.ERole) (err error) {
	err = pcvSetDefaultEndpoint(v, deviceID, eRole)
	return
}

func pcvSetDefaultEndpoint(pcv *IPolicyConfigVista, deviceID string, eRole wca.ERole) (err error) {
	var ptr *uint16
	if ptr, err = syscall.UTF16PtrFromString(deviceID); err != nil {
		return
	}
	hr, _, _ := syscall.Syscall(
		pcv.VTable().SetDefaultEndpoint,
		3,
		uintptr(unsafe.Pointer(pcv)),
		uintptr(unsafe.Pointer(ptr)),
		uintptr(uint32(eRole)))
	if hr != 0 {
		err = ole.NewError(hr)
	}
	return
}

func setAudioDeviceByID(deviceID string) {
	GUID_IPolicyConfigVista := ole.NewGUID("{568b9108-44bf-40b4-9006-86afe5b5a620}")
	GUID_CPolicyConfigVistaClient := ole.NewGUID("{294935CE-F637-4E7C-A41B-AB255460B862}")
	var policyConfig *IPolicyConfigVista

	if err := wca.CoCreateInstance(
		GUID_CPolicyConfigVistaClient,
		0,
		wca.CLSCTX_ALL,
		GUID_IPolicyConfigVista,
		&policyConfig,
	); err != nil {
		panic(err)
	}
	defer policyConfig.Release()

	if err := policyConfig.SetDefaultEndpoint(deviceID, wca.EConsole); err != nil {
		panic(err)
	}

	deviceName := getDeviceName(deviceID)

	err := beeep.Notify(
		"Audio device changed",
		deviceName,
		"assets/audio-speaker.png",
	)
	if err != nil {
		panic(err)
	}
}

func printCurrentDevice() {
	var mmde *wca.IMMDeviceEnumerator
	if err := wca.CoCreateInstance(wca.CLSID_MMDeviceEnumerator, 0, wca.CLSCTX_ALL, wca.IID_IMMDeviceEnumerator, &mmde); err != nil {
		return
	}
	defer mmde.Release()

	var mmd *wca.IMMDevice
	if err := mmde.GetDefaultAudioEndpoint(wca.ERender, wca.EConsole, &mmd); err != nil {
		return
	}
	defer mmd.Release()

	var ps *wca.IPropertyStore
	if err := mmd.OpenPropertyStore(wca.STGM_READ, &ps); err != nil {
		return
	}
	defer ps.Release()

	var pv wca.PROPVARIANT
	if err := ps.GetValue(&wca.PKEY_Device_FriendlyName, &pv); err != nil {
		return
	}

	var deviceName = pv.String()

	fmt.Printf("%s", deviceName)

	err := beeep.Notify(
		"Current audio device",
		deviceName,
		"assets/audio-speaker.png",
	)
	if err != nil {
		panic(err)
	}
}
