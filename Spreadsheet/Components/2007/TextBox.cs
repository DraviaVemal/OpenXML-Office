// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Global_2007;

namespace OpenXMLOffice.Spreadsheet_2007
{

	/// <summary>
	/// Represents Text box class to build on
	/// </summary>
	public class TextBox : SpreadSheetCommonProperties
	{
		private readonly TextBoxSetting textBoxSetting;
		// private readonly XDR.Shape documentShape;
		private readonly Worksheet worksheet;
		/// <summary>
		/// Create Text box with provided settings
		/// </summary>
		internal TextBox(TextBoxSetting TextBoxSetting)
		{
			textBoxSetting = TextBoxSetting;
		}
		/// <summary>
		/// Create Text box with provided settings
		/// </summary>
		public TextBox(Worksheet Worksheet, TextBoxSetting TextBoxSetting)
		{
			worksheet = Worksheet;
			textBoxSetting = TextBoxSetting;
		}
		// /// <summary>
		// /// Get Text box Shape
		// /// </summary>
		// internal XDR.Shape GetTextBoxShape()
		// {
		// 	return documentShape;
		// }
	}
}
