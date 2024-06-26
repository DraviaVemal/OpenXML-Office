// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
namespace OpenXMLOffice.Presentation_2007
{
	/// <summary>
	/// PPT View Properties Part Handling
	/// </summary>
	public class ViewProperties
	{
		private readonly P.ViewProperties documentViewProperties = new P.ViewProperties();
		/// <summary>
		/// Create New View Properties
		/// </summary>
		public ViewProperties()
		{
			documentViewProperties.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
			documentViewProperties.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
			documentViewProperties.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");
			documentViewProperties.AppendChild(CreateNormalViewProperties());
			documentViewProperties.AppendChild(CreateSlideViewProperties());
			documentViewProperties.AppendChild(CreateNotesTextViewProperties());
			documentViewProperties.AppendChild(new P.GridSpacing()
			{
				Cx = 72008,
				Cy = 72008
			});
		}
		/// <summary>
		/// Return OpenXML View Properties
		/// </summary>
		/// <returns>
		/// </returns>
		public P.ViewProperties GetViewProperties()
		{
			return documentViewProperties;
		}
		private static P.NotesTextViewProperties CreateNotesTextViewProperties()
		{
			P.NotesTextViewProperties notesTextViewProperties = new P.NotesTextViewProperties(
				new P.CommonViewProperties
				{
					ScaleFactor = new P.ScaleFactor(new A.ScaleX()
					{
						Denominator = 1,
						Numerator = 1
					}, new A.ScaleY()
					{
						Denominator = 1,
						Numerator = 1
					}),
					Origin = new P.Origin()
					{
						X = 0,
						Y = 0,
					},
				}
			);
			return notesTextViewProperties;
		}
		private static P.NormalViewProperties CreateNormalViewProperties()
		{
			P.NormalViewProperties normalViewProperties = new P.NormalViewProperties()
			{
				HorizontalBarState = P.SplitterBarStateValues.Maximized,
				RestoredLeft = new P.RestoredLeft { AutoAdjust = false, Size = 16014 },
				RestoredTop = new P.RestoredTop { Size = 94660 }
			};
			return normalViewProperties;
		}
		private static P.SlideViewProperties CreateSlideViewProperties()
		{
			P.SlideViewProperties slideViewProperties = new P.SlideViewProperties(
				new P.CommonSlideViewProperties(
					new P.CommonViewProperties
					{
						VariableScale = true,
						ScaleFactor = new P.ScaleFactor(new A.ScaleX()
						{
							Denominator = 100,
							Numerator = 159
						}, new A.ScaleY()
						{
							Denominator = 100,
							Numerator = 159
						}),
						Origin = new P.Origin()
						{
							X = 306,
							Y = 138,
						},
					},
					new P.GuideList()
				)
				{
					SnapToGrid = false
				}
			);
			return slideViewProperties;
		}
	}
}
