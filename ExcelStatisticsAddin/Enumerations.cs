namespace VS.NET_RefeditControl
{
	/// <summary>
	/// The cell reference type for the control
	/// </summary>
	public enum ReferenceStyle
	{
		/// <summary>
		/// Cell references are Absolute ($A$1)
		/// </summary>
		Absolute = 0,
		/// <summary>
		/// Cell references are Row relative ($A1)
		/// </summary>
		RowRelative = 1,
		/// <summary>
		/// Cell references are Column relative (A$1)
		/// </summary>
		ColumnRelative = 2,
		/// <summary>
		/// Cell references are Relative (A1)
		/// </summary>
		Relative = 3,
	}
	/// <summary>
	/// The display style to use when the the control is collapsed
	/// </summary>
	public enum CollapseStyle
	{
		/// <summary>
		/// No Change
		/// </summary>
		None = 0,
		/// <summary>
		/// The container Form is collapsed and the control resized to fill the width of the form (same as Excel's default behaviour).
		/// </summary>
		CollapseFormAndFitCellSelector = 1,
		/// <summary>
		/// The container Form is collapsed but the control is not resized.
		/// </summary>
		CollapseFormOnly = 2,
	}
	/// <summary>
	/// The display style for the selected cell reference
	/// </summary>
	public enum DisplayStyle
	{
		/// <summary>
		/// The cell's address only is used 
		/// <example>$A$1</example>
		/// </summary>
		AddressRange = 0,
		/// <summary>
		/// The calculated SUM of each cell is used 
		/// <example>If $A$1 = 100 and $A$2 = 50 and both cells are selected, the display will show 150</example>
		/// </summary>
		SumRangeValues = 1,
		/// <summary>
		/// Each cell's address is concatenated together using <see cref="ResultSeparator"/> 
		/// <example>$A$1</example>
		/// </summary>
		ConcatenateRangeValues = 2,
		/// <summary>
		/// The worksheet name is appended to the fron of the cell address
		/// <example>Sheet1!$A$1</example>
		/// </summary>
		AddressRangeWithSheet = 3,
		/// <summary>
		/// The workbook and worksheet names are appended to the fron of the cell address
		/// <example>[Book1]Sheet1!$A$1</example>
		/// </summary>
		AddressRangeWithSheetBook = 4,
	}
	/// <summary>
	/// When results are concatenated using <see cref="ResultSeparator"/>.ConcatenateRangeValues, the separator string to use 
	/// </summary>
	public enum ResultSeparator
	{
		/// <summary>
		/// No separator
		/// </summary>
		/// <example>$A$1$A$2</example>
		None = 0,
		/// <summary>
		/// Uses a single "space" character
		/// </summary>
		/// <example>$A$1 $A$2</example>
		Space = 1,
		/// <summary>
		/// Uses a semi-colon
		/// <example>$A$1;$A$2</example>
		/// </summary>
		SemiColon = 2,
		/// <summary>
		/// Uses a colon
		/// <example>$A$1:$A$2</example>
		/// </summary>
		Colon = 3,
		/// <summary>
		/// Uses a pipe (|) character
		/// <example>$A$1|$A$2</example>
		/// </summary>
		Pipe = 4,
		/// <summary>
		/// Uses a comma
		/// <example>$A$1,$A$2</example>
		/// </summary>
		Comma = 5,
	}
}
