namespace NetOdt
{
    /// <summary>
    /// The type of a value inside a table cell
    /// </summary>
    public enum CellValueType
    {
        /// <summary>
        /// A decimal value
        /// </summary>
        Float,

        /// <summary>
        /// A percentage value
        /// </summary>
        Percentage,

        /// <summary>
        /// A currency value
        /// </summary>
        Currency,

        /// <summary>
        /// A date value
        /// </summary>
        Date,

        /// <summary>
        /// A time value
        /// </summary>
        Time,

        /// <summary>
        /// A boolean value
        /// </summary>
        Boolean,

        /// <summary>
        /// A text
        /// </summary>
        String,
    }
}
