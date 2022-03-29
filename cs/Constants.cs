using System;
using Windows.Devices.Bluetooth.GenericAttributeProfile;

namespace SDKTemplate
{
    // Define the characteristics and other properties of our custom service.
    public class Constants
    {
        // BT_Code: Initializes custom local parameters w/ properties, protection levels as well as common descriptors like User Description. 
        public static readonly GattLocalCharacteristicParameters gattOperandParameters = new GattLocalCharacteristicParameters
        {
            CharacteristicProperties = GattCharacteristicProperties.Write |
                                       GattCharacteristicProperties.WriteWithoutResponse,
            WriteProtectionLevel = GattProtectionLevel.Plain,
            UserDescription = "Operand Characteristic"
        };

        public static readonly GattLocalCharacteristicParameters gattOperatorParameters = new GattLocalCharacteristicParameters
        {
            CharacteristicProperties = GattCharacteristicProperties.Write |
                                       GattCharacteristicProperties.WriteWithoutResponse,
            WriteProtectionLevel = GattProtectionLevel.Plain,
            UserDescription = "Operator Characteristic"
        };

        public static readonly GattLocalCharacteristicParameters gattResultParameters = new GattLocalCharacteristicParameters
        {
            CharacteristicProperties = GattCharacteristicProperties.Read |
                                       GattCharacteristicProperties.Notify,
            WriteProtectionLevel = GattProtectionLevel.Plain,
            UserDescription = "Result Characteristic"
        };

        public static readonly Guid CalcServiceUuid = Guid.Parse("caecface-e1d9-11e6-bf01-fe55135034f0");

        public static readonly Guid Op1CharacteristicUuid = Guid.Parse("caec2ebc-e1d9-11e6-bf01-fe55135034f1");
        public static readonly Guid Op2CharacteristicUuid = Guid.Parse("caec2ebc-e1d9-11e6-bf01-fe55135034f2");
        public static readonly Guid OperatorCharacteristicUuid = Guid.Parse("caec2ebc-e1d9-11e6-bf01-fe55135034f3");
        public static readonly Guid ResultCharacteristicUuid = Guid.Parse("caec2ebc-e1d9-11e6-bf01-fe55135034f4");

        // TODO: Define new constants for RemoteHotkeyService/Char
        public static readonly Guid RemoteHotkeyServiceUuid = Guid.Parse("0000dd10-0000-1000-8000-00805f9b34fb");
        public static readonly Guid RemoteHotkeyCharacteristicUuid = Guid.Parse("0000dd11-0000-1000-8000-00805f9b34fb");
    };
}
