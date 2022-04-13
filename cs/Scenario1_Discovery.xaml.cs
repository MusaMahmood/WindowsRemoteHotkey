//*********************************************************
//
// Copyright (c) Microsoft. All rights reserved.
// This code is licensed under the MIT License (MIT).
// THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY
// IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR
// PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.
//
//*********************************************************

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Threading.Tasks;
using Windows.Devices.Bluetooth;
using Windows.Devices.Bluetooth.GenericAttributeProfile;
using Windows.Devices.Enumeration;
using Windows.Security.Cryptography;
using Windows.System;
using Windows.UI.Core;
using Windows.UI.Input.Preview.Injection;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Navigation;

namespace SDKTemplate
{
    // This scenario uses a DeviceWatcher to enumerate nearby Bluetooth Low Energy devices,
    // displays them in a ListView, and lets the user select a device and pair it.
    // This device will be used by future scenarios.
    // For more information about device discovery and pairing, including examples of
    // customizing the pairing process, see the DeviceEnumerationAndPairing sample.
    public sealed partial class Scenario1_Discovery : Page
    {
        private MainPage rootPage = MainPage.Current;

        private ObservableCollection<BluetoothLEDeviceDisplay> KnownDevices = new ObservableCollection<BluetoothLEDeviceDisplay>();
        private List<DeviceInformation> UnknownDevices = new List<DeviceInformation>();

        private DeviceWatcher deviceWatcher;

        private BluetoothLEDevice bluetoothLeDevice = null;
        private GattCharacteristic selectedCharacteristic;

        // Only one registered characteristic at a time.
        private GattCharacteristic registeredCharacteristic;

        #region Error Codes
        readonly int E_BLUETOOTH_ATT_WRITE_NOT_PERMITTED = unchecked((int)0x80650003);
        readonly int E_BLUETOOTH_ATT_INVALID_PDU = unchecked((int)0x80650004);
        readonly int E_ACCESSDENIED = unchecked((int)0x80070005);
        readonly int E_DEVICE_NOT_AVAILABLE = unchecked((int)0x800710df); // HRESULT_FROM_WIN32(ERROR_DEVICE_NOT_AVAILABLE)
        #endregion

        #region UI Code
        public Scenario1_Discovery()
        {
            InitializeComponent();
        }

        protected override void OnNavigatedTo(NavigationEventArgs e) {
           if (deviceWatcher == null) {
                StartBleDeviceWatcher();
                EnumerateButton.Content = "Stop scan";
                rootPage.NotifyUser($"Device watcher started", NotifyType.StatusMessage);
            } else {
                StopBleDeviceWatcher();
                EnumerateButton.Content = "Start enumerating";
                rootPage.NotifyUser($"Device watcher stopped.", NotifyType.StatusMessage);
            }
        }

        protected override async void OnNavigatedFrom(NavigationEventArgs e)
        {
            StopBleDeviceWatcher();

            // Save the selected device's ID for use in other scenarios.
            var bleDeviceDisplay = ResultsListView.SelectedItem as BluetoothLEDeviceDisplay;
            if (bleDeviceDisplay != null)
            {
                rootPage.SelectedBleDeviceId = bleDeviceDisplay.Id;
                rootPage.SelectedBleDeviceName = bleDeviceDisplay.Name;
            }

            var success = await ClearBluetoothLEDeviceAsync();
            if (!success) {
                rootPage.NotifyUser("Error: Unable to reset app state", NotifyType.ErrorMessage);
            }
        }

        private async Task<bool> ClearBluetoothLEDeviceAsync() {
            if (subscribedForNotifications) {
                // Need to clear the CCCD from the remote device so we stop receiving notifications
                var result = await registeredCharacteristic.WriteClientCharacteristicConfigurationDescriptorAsync(GattClientCharacteristicConfigurationDescriptorValue.None);
                if (result != GattCommunicationStatus.Success) {
                    return false;
                }
                else {
                    selectedCharacteristic.ValueChanged -= Characteristic_ValueChanged;
                    subscribedForNotifications = false;
                }
            }
            bluetoothLeDevice?.Dispose();
            bluetoothLeDevice = null;
            return true;
        }

        private void EnumerateButton_Click()
        {
            if (deviceWatcher == null)
            {
                StartBleDeviceWatcher();
                EnumerateButton.Content = "Stop enumerating";
                rootPage.NotifyUser($"Device watcher started.", NotifyType.StatusMessage);
            }
            else
            {
                StopBleDeviceWatcher();
                EnumerateButton.Content = "Start enumerating";
                rootPage.NotifyUser($"Device watcher stopped.", NotifyType.StatusMessage);
            }
        }

        private bool Not(bool value) => !value;

        #endregion

        #region Device discovery

        /// <summary>
        /// Starts a device watcher that looks for all nearby Bluetooth devices (paired or unpaired). 
        /// Attaches event handlers to populate the device collection.
        /// </summary>
        private void StartBleDeviceWatcher()
        {
            // Additional properties we would like about the device.
            // Property strings are documented here https://msdn.microsoft.com/en-us/library/windows/desktop/ff521659(v=vs.85).aspx
            string[] requestedProperties = { "System.Devices.Aep.DeviceAddress", "System.Devices.Aep.IsConnected", "System.Devices.Aep.Bluetooth.Le.IsConnectable" };

            // BT_Code: Example showing paired and non-paired in a single query.
            string aqsAllBluetoothLEDevices = "(System.Devices.Aep.ProtocolId:=\"{bb7bb05e-5972-42b5-94fc-76eaa7084d49}\")";

            deviceWatcher =
                    DeviceInformation.CreateWatcher(
                        aqsAllBluetoothLEDevices,
                        requestedProperties,
                        DeviceInformationKind.AssociationEndpoint);

            // Register event handlers before starting the watcher.
            deviceWatcher.Added += DeviceWatcher_Added;
            deviceWatcher.Updated += DeviceWatcher_Updated;
            deviceWatcher.Removed += DeviceWatcher_Removed;
            deviceWatcher.EnumerationCompleted += DeviceWatcher_EnumerationCompleted;
            deviceWatcher.Stopped += DeviceWatcher_Stopped;

            // Start over with an empty collection.
            KnownDevices.Clear();

            // Start the watcher. Active enumeration is limited to approximately 30 seconds.
            // This limits power usage and reduces interference with other Bluetooth activities.
            // To monitor for the presence of Bluetooth LE devices for an extended period,
            // use the BluetoothLEAdvertisementWatcher runtime class. See the BluetoothAdvertisement
            // sample for an example.
            deviceWatcher.Start();
        }

        /// <summary>
        /// Stops watching for all nearby Bluetooth devices.
        /// </summary>
        private void StopBleDeviceWatcher()
        {
            if (deviceWatcher != null)
            {
                // Unregister the event handlers.
                deviceWatcher.Added -= DeviceWatcher_Added;
                deviceWatcher.Updated -= DeviceWatcher_Updated;
                deviceWatcher.Removed -= DeviceWatcher_Removed;
                deviceWatcher.EnumerationCompleted -= DeviceWatcher_EnumerationCompleted;
                deviceWatcher.Stopped -= DeviceWatcher_Stopped;

                // Stop the watcher.
                deviceWatcher.Stop();
                deviceWatcher = null;
            }
        }

        private BluetoothLEDeviceDisplay FindBluetoothLEDeviceDisplay(string id)
        {
            foreach (BluetoothLEDeviceDisplay bleDeviceDisplay in KnownDevices)
            {
                if (bleDeviceDisplay.Id == id)
                {
                    return bleDeviceDisplay;
                }
            }
            return null;
        }

        private DeviceInformation FindUnknownDevices(string id)
        {
            foreach (DeviceInformation bleDeviceInfo in UnknownDevices)
            {
                if (bleDeviceInfo.Id == id)
                {
                    return bleDeviceInfo;
                }
            }
            return null;
        }

        private async void DeviceWatcher_Added(DeviceWatcher sender, DeviceInformation deviceInfo)
        {
            // We must update the collection on the UI thread because the collection is databound to a UI element.
            await Dispatcher.RunAsync(CoreDispatcherPriority.Normal, () =>
            {
                lock (this)
                {
                    Debug.WriteLine(String.Format("Added {0}{1}", deviceInfo.Id, deviceInfo.Name));

                    // Protect against race condition if the task runs after the app stopped the deviceWatcher.
                    if (sender == deviceWatcher)
                    {
                        // Make sure device isn't already present in the list.
                        if (FindBluetoothLEDeviceDisplay(deviceInfo.Id) == null)
                        {
                            if (deviceInfo.Name != string.Empty)
                            {
                                //rootPage.NotifyUser(String.Format("Added [{0}]   [{1}]", deviceInfo.Id, deviceInfo.Name), NotifyType.StatusMessage);
                                if (deviceInfo.Id.Contains("7e:19:d5:19:f9:6c")) {
                                    rootPage.NotifyUser(String.Format("Paired Device Detected. [Name: {0}, MAC addr: 7e:19:d5:19:f9:6c]", deviceInfo.Name), NotifyType.StatusMessage);
                                    rootPage.SelectedBleDeviceId = deviceInfo.Id;
                                    rootPage.SelectedBleDeviceName = deviceInfo.Name;
                                    //TODO: Launch connection & enable notifs:
                                    StopBleDeviceWatcher();
                                    EnumerateButton.Content = "Start enumerating";
                                    rootPage.NotifyUser($"Device watcher stopped.", NotifyType.StatusMessage);
                                    ConnectAndEnableDevice();
                                }
                                // If device has a friendly name display it immediately.
                                KnownDevices.Add(new BluetoothLEDeviceDisplay(deviceInfo));
                            }
                            else
                            {
                                // Add it to a list in case the name gets updated later. 
                                UnknownDevices.Add(deviceInfo);
                            }
                        }

                    }
                }
            });
        }

        private async void ConnectAndEnableDevice() {
            try {
                // BT_Code: BluetoothLEDevice.FromIdAsync must be called from a UI thread because it may prompt for consent.
                bluetoothLeDevice = await BluetoothLEDevice.FromIdAsync(rootPage.SelectedBleDeviceId);

                if (bluetoothLeDevice == null) {
                    rootPage.NotifyUser("Failed to connect to device.", NotifyType.ErrorMessage);
                }
            }
            catch (Exception ex) when (ex.HResult == E_DEVICE_NOT_AVAILABLE) {
                rootPage.NotifyUser("Bluetooth radio is not on.", NotifyType.ErrorMessage);
            }

            if (bluetoothLeDevice != null) {
                // Note: BluetoothLEDevice.GattServices property will return an empty list for unpaired devices. For all uses we recommend using the GetGattServicesAsync method.
                // BT_Code: GetGattServicesAsync returns a list of all the supported services of the device (even if it's not paired to the system).
                // If the services supported by the device are expected to change during BT usage, subscribe to the GattServicesChanged event.
                GattDeviceServicesResult result = await bluetoothLeDevice.GetGattServicesAsync(BluetoothCacheMode.Uncached);

                IReadOnlyList<GattCharacteristic> characteristics = null;

                if (result.Status == GattCommunicationStatus.Success) {
                    var services = result.Services;
                    //rootPage.NotifyUser(String.Format("Found {0} services", services.Count), NotifyType.StatusMessage);
                    foreach (var service in services) {
                        //ServiceList.Items.Add(new ComboBoxItem { Content = DisplayHelpers.GetServiceName(service), Tag = service });
                        //rootPage.NotifyUser(String.Format("{0}", service.Uuid.ToString()), NotifyType.StatusMessage);
                        if (service.Uuid.ToString().Contains("0000dd10-0000-1000-8000-00805f9b34fb")) {
                            Debug.WriteLine(String.Format("New Service UUID {0}", service.Uuid.ToString()));
                            // TODO: enumerate characteristics: 
                            try {
                                // Ensure we have access to the device.
                                var accessStatus = await service.RequestAccessAsync();
                                if (accessStatus == DeviceAccessStatus.Allowed) {
                                    // BT_Code: Get all the child characteristics of a service. Use the cache mode to specify uncached characterstics only 
                                    // and the new Async functions to get the characteristics of unpaired devices as well. 
                                    var result2 = await service.GetCharacteristicsAsync(BluetoothCacheMode.Uncached);
                                    if (result2.Status == GattCommunicationStatus.Success) {
                                        characteristics = result2.Characteristics;
                                    }
                                    else {
                                        rootPage.NotifyUser("Error accessing service.", NotifyType.ErrorMessage);

                                        // On error, act as if there are no characteristics.
                                        characteristics = new List<GattCharacteristic>();
                                    }
                                }
                                else {
                                    // Not granted access
                                    rootPage.NotifyUser("Error accessing service.", NotifyType.ErrorMessage);

                                    // On error, act as if there are no characteristics.
                                    characteristics = new List<GattCharacteristic>();

                                }
                            }
                            catch (Exception ex) {
                                rootPage.NotifyUser("Restricted service. Can't read characteristics: " + ex.Message, NotifyType.ErrorMessage);
                                // On error, act as if there are no characteristics.
                                characteristics = new List<GattCharacteristic>();
                            }
                            // Check characteristics for string "dd11"
                            foreach (GattCharacteristic c in characteristics) {
                                if (c.Uuid.ToString().Contains("dd11")) {
                                    Debug.WriteLine(String.Format("Found Characteristic UUID {0}", c.Uuid.ToString()));
                                    rootPage.NotifyUser(String.Format("Found Characteristic UUID {0}", c.Uuid.ToString()), NotifyType.StatusMessage);
                                    // TODO: enable characteristic notifications
                                    selectedCharacteristic = c;
                                    EnableCharacteristicNotifications();
                                }
                            }
                        }
                    }
                }
                else {
                    rootPage.NotifyUser("Device unreachable", NotifyType.ErrorMessage);
                }
            }                
        }

        private bool subscribedForNotifications = false;

        private async void EnableCharacteristicNotifications() {
            // Get all the child descriptors of a characteristics. Use the cache mode to specify uncached descriptors only 
            // and the new Async functions to get the descriptors of unpaired devices as well. 
            var result = await selectedCharacteristic.GetDescriptorsAsync(BluetoothCacheMode.Uncached);
            if (result.Status != GattCommunicationStatus.Success) {
                rootPage.NotifyUser("Descriptor read failure: " + result.Status.ToString(), NotifyType.ErrorMessage);
                return;
            }
            // initialize status
            GattCommunicationStatus status = GattCommunicationStatus.Unreachable;
            var cccdValue = GattClientCharacteristicConfigurationDescriptorValue.None;
            if (selectedCharacteristic.CharacteristicProperties.HasFlag(GattCharacteristicProperties.Indicate)) {
                cccdValue = GattClientCharacteristicConfigurationDescriptorValue.Indicate;
            } else if (selectedCharacteristic.CharacteristicProperties.HasFlag(GattCharacteristicProperties.Notify)) {
                cccdValue = GattClientCharacteristicConfigurationDescriptorValue.Notify;
            }

            try {
                // BT_Code: Must write the CCCD in order for server to send indications.
                // We receive them in the ValueChanged event handler.
                status = await selectedCharacteristic.WriteClientCharacteristicConfigurationDescriptorAsync(cccdValue);

                if (status == GattCommunicationStatus.Success) {
                    AddValueChangedHandler();
                    rootPage.NotifyUser("Successfully subscribed for value changes", NotifyType.StatusMessage);
                }
                else {
                    rootPage.NotifyUser($"Error registering for value changes: {status}", NotifyType.ErrorMessage);
                }
            }
            catch (UnauthorizedAccessException ex) {
                // This usually happens when a device reports that it support indicate, but it actually doesn't.
                rootPage.NotifyUser(ex.Message, NotifyType.ErrorMessage);
            }
        }

        private void AddValueChangedHandler() {
            //ValueChangedSubscribeToggle.Content = "Unsubscribe from value changes";
            if (!subscribedForNotifications) {
                registeredCharacteristic = selectedCharacteristic;
                registeredCharacteristic.ValueChanged += Characteristic_ValueChanged;
                subscribedForNotifications = true;
            }
        }

        private async void Characteristic_ValueChanged(GattCharacteristic sender, GattValueChangedEventArgs args) {
            // BT_Code: An Indicate or Notify reported that the value has changed.
            // Display the new value with a timestamp.

            byte[] data;
            CryptographicBuffer.CopyToByteArray(args.CharacteristicValue, out data);

            ParseData(data);

            var newValue = ConvertByteArrayToString(data);
            var message = $"Value at {DateTime.Now:hh:mm:ss.FFF}: {newValue}";
            await Dispatcher.RunAsync(CoreDispatcherPriority.Normal, () => CharacteristicLatestValue.Text = message);
        }

        private string ConvertByteArrayToString(byte[] data) {
            return BitConverter.ToString(data);
        }

        private void ParseData(byte[] data) {
            // Parse First Byte: [Ctrl, Shift, Alt, Win | 0010] < last four bits irrelevant
            bool[] bArray = new bool[4];
            bArray[0] = ((uint)data[0] & 0b_1000_0000) == 0b_1000_0000; // Ctrl
            bArray[1] = ((uint)data[0] & 0b_0100_0000) == 0b_0100_0000; // Shift
            bArray[2] = ((uint)data[0] & 0b_0010_0000) == 0b_0010_0000; // Alt
            bArray[3] = ((uint)data[0] & 0b_0001_0000) == 0b_0001_0000; // WinKey
            rootPage.NotifyUser($"byteMatch: {bArray[0]} {bArray[1]} {bArray[2]} {bArray[3]} | Vkey: {(ushort)data[1]}", NotifyType.StatusMessage);
            // Parse Second Byte
            Hotkey_Click(bArray, (ushort)data[1]);
        }

        //TODO: Requst commands (ushort, up to 4), make unused null if not required. 
        private async void Hotkey_Click(bool[] bArray, ushort virtualKey) {
            InputInjector inputInjector = InputInjector.TryCreate();

            var ctrl = new InjectedInputKeyboardInfo();
            ctrl.VirtualKey = (ushort)VirtualKey.Control;
            ctrl.KeyOptions = InjectedInputKeyOptions.None; // ExtendedKey?

            var shift = new InjectedInputKeyboardInfo();
            shift.VirtualKey = (ushort)VirtualKey.Shift;
            shift.KeyOptions = InjectedInputKeyOptions.None;

            var alt = new InjectedInputKeyboardInfo();
            alt.VirtualKey = (ushort)VirtualKey.Menu;
            alt.KeyOptions = InjectedInputKeyOptions.None;

            var win = new InjectedInputKeyboardInfo();
            win.VirtualKey = (ushort)VirtualKey.LeftWindows;
            win.KeyOptions = InjectedInputKeyOptions.None;

            var keyCommand = new InjectedInputKeyboardInfo();
            keyCommand.VirtualKey = virtualKey;
            keyCommand.KeyOptions = InjectedInputKeyOptions.None;

            List<InjectedInputKeyboardInfo> keyInjectList = new List<InjectedInputKeyboardInfo>();
            if (bArray[0]) {
                keyInjectList.Add(ctrl);
            }
            if (bArray[1]) {
                keyInjectList.Add(shift);
            }
            if (bArray[3]) {
                keyInjectList.Add(win);
            }
            if (bArray[2]) {
                keyInjectList.Add(alt);
            }
            if (virtualKey >= 0) {
                keyInjectList.Add(keyCommand);
            }

            // TODO: Ensure ctrl, shift, win, alt keys are all released before pressing?

            // Sets key down 
            inputInjector.InjectKeyboardInput(keyInjectList);

            // release all keys after use
            ctrl.KeyOptions = InjectedInputKeyOptions.KeyUp;
            shift.KeyOptions = InjectedInputKeyOptions.KeyUp;
            alt.KeyOptions = InjectedInputKeyOptions.KeyUp;
            win.KeyOptions = InjectedInputKeyOptions.KeyUp;
            keyCommand.KeyOptions = InjectedInputKeyOptions.KeyUp;

            inputInjector.InjectKeyboardInput(new[] { win, alt, ctrl, shift, keyCommand });
        }

        private async void DeviceWatcher_Updated(DeviceWatcher sender, DeviceInformationUpdate deviceInfoUpdate)
        {
            // We must update the collection on the UI thread because the collection is databound to a UI element.
            await Dispatcher.RunAsync(CoreDispatcherPriority.Normal, () =>
            {
                lock (this)
                {
                    Debug.WriteLine(String.Format("Updated {0}{1}", deviceInfoUpdate.Id, ""));

                    // Protect against race condition if the task runs after the app stopped the deviceWatcher.
                    if (sender == deviceWatcher)
                    {
                        BluetoothLEDeviceDisplay bleDeviceDisplay = FindBluetoothLEDeviceDisplay(deviceInfoUpdate.Id);
                        if (bleDeviceDisplay != null)
                        {
                            // Device is already being displayed - update UX.
                            bleDeviceDisplay.Update(deviceInfoUpdate);
                            return;
                        }

                        DeviceInformation deviceInfo = FindUnknownDevices(deviceInfoUpdate.Id);
                        if (deviceInfo != null)
                        {
                            deviceInfo.Update(deviceInfoUpdate);
                            // If device has been updated with a friendly name it's no longer unknown.
                            if (deviceInfo.Name != String.Empty)
                            {
                                KnownDevices.Add(new BluetoothLEDeviceDisplay(deviceInfo));
                                UnknownDevices.Remove(deviceInfo);
                            }
                        }
                    }
                }
            });
        }

        private async void DeviceWatcher_Removed(DeviceWatcher sender, DeviceInformationUpdate deviceInfoUpdate)
        {
            // We must update the collection on the UI thread because the collection is databound to a UI element.
            await Dispatcher.RunAsync(CoreDispatcherPriority.Normal, () =>
            {
                lock (this)
                {
                    Debug.WriteLine(String.Format("Removed {0}{1}", deviceInfoUpdate.Id,""));

                    // Protect against race condition if the task runs after the app stopped the deviceWatcher.
                    if (sender == deviceWatcher)
                    {
                        // Find the corresponding DeviceInformation in the collection and remove it.
                        BluetoothLEDeviceDisplay bleDeviceDisplay = FindBluetoothLEDeviceDisplay(deviceInfoUpdate.Id);
                        if (bleDeviceDisplay != null)
                        {
                            KnownDevices.Remove(bleDeviceDisplay);
                        }

                        DeviceInformation deviceInfo = FindUnknownDevices(deviceInfoUpdate.Id);
                        if (deviceInfo != null)
                        {
                            UnknownDevices.Remove(deviceInfo);
                        }
                    }
                }
            });
        }

        private async void DeviceWatcher_EnumerationCompleted(DeviceWatcher sender, object e)
        {
            // We must update the collection on the UI thread because the collection is databound to a UI element.
            await Dispatcher.RunAsync(CoreDispatcherPriority.Normal, () =>
            {
                // Protect against race condition if the task runs after the app stopped the deviceWatcher.
                if (sender == deviceWatcher)
                {
                    //rootPage.NotifyUser($"{KnownDevices.Count} devices found. Enumeration completed.", NotifyType.StatusMessage);
                }
            });
        }

        private async void DeviceWatcher_Stopped(DeviceWatcher sender, object e)
        {
            // We must update the collection on the UI thread because the collection is databound to a UI element.
            await Dispatcher.RunAsync(CoreDispatcherPriority.Normal, () =>
            {
                // Protect against race condition if the task runs after the app stopped the deviceWatcher.
                if (sender == deviceWatcher)
                {
                    //rootPage.NotifyUser($"No longer watching for devices.", sender.Status == DeviceWatcherStatus.Aborted ? NotifyType.ErrorMessage : NotifyType.StatusMessage);
                }
            });
        }
        #endregion

        #region Pairing

        private bool isBusy = false;

        private async void PairButton_Click()
        {
            // Do not allow a new Pair operation to start if an existing one is in progress.
            if (isBusy)
            {
                return;
            }

            isBusy = true;

            rootPage.NotifyUser("Pairing started. Please wait...", NotifyType.StatusMessage);

            // For more information about device pairing, including examples of
            // customizing the pairing process, see the DeviceEnumerationAndPairing sample.

            // Capture the current selected item in case the user changes it while we are pairing.
            var bleDeviceDisplay = ResultsListView.SelectedItem as BluetoothLEDeviceDisplay;

            // BT_Code: Pair the currently selected device.
            DevicePairingResult result = await bleDeviceDisplay.DeviceInformation.Pairing.PairAsync();
            rootPage.NotifyUser($"Pairing result = {result.Status}",
                result.Status == DevicePairingResultStatus.Paired || result.Status == DevicePairingResultStatus.AlreadyPaired
                    ? NotifyType.StatusMessage
                    : NotifyType.ErrorMessage);

            isBusy = false;
        }

        #endregion
    }
}