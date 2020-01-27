namespace NetOdt
{
    /// <summary>
    /// The type of a value
    /// </summary>
    public enum OfficeValueType
    {
        /// <summary>
        /// A decimal value (e.g. 1, 2.2)
        /// </summary>
        Float,

        /// <summary>
        /// A percentage value (e.g. 1%, 2.20 %)
        /// </summary>
        Percentage,

        /// <summary>
        /// A currency value (e.g. 1 €, 2.20 €)
        /// </summary>
        Currency,

        /// <summary>
        /// A date value (e.g. 01.01.2010, 22.12.2020)
        /// </summary>
        Date,

        /// <summary>
        /// A time value (e.g. 01:01:01, 22:22:22)
        /// </summary>
        Time,

        /// <summary>
        /// A scientific value (e.g. 1.11E+11, 2.22E+22)
        /// </summary>
        Scientific,

        /// <summary>
        /// A fraction value (e.g. 1 1/2, 2 1/2)
        /// </summary>
        Fraction,

        /// <summary>
        /// A boolean value (true or false)
        /// </summary>
        Boolean,

        /// <summary>
        /// A text
        /// </summary>
        String,
    }
}
