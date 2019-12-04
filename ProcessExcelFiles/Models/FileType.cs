namespace ProcessExcelFiles
{
    public enum FileType
    {
        Boolean = 0,
        //
        // Summary:
        //     Number.
        //     When the item is serialized out as xml, its value is "n".
        Number = 1,
        //
        // Summary:
        //     Error.
        //     When the item is serialized out as xml, its value is "e".
        Error = 2,
        //
        // Summary:
        //     Shared String.
        //     When the item is serialized out as xml, its value is "s".
        SharedString = 3,
        //
        // Summary:
        //     String.
        //     When the item is serialized out as xml, its value is "str".
        String = 4,
        //
        // Summary:
        //     Inline String.
        //     When the item is serialized out as xml, its value is "inlineStr".
        InlineString = 5,
        //
        // Summary:
        //     d.
        //     When the item is serialized out as xml, its value is "d".
        Date = 6
    }
}