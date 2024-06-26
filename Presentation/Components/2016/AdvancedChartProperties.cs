// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
using OpenXMLOffice.Global_2007;
using OpenXMLOffice.Presentation_2007;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using CX = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using System.Collections.Generic;
namespace OpenXMLOffice.Presentation_2016
{
	/// <summary>
	///
	/// </summary>
	public class AdvancedChartProperties : ChartProperties
	{
		private AlternateContent alternateContent;
		private readonly TextBox errorMessage;
		/// <summary>
		///
		/// </summary>
		public AdvancedChartProperties(Slide slide, ChartSetting<PresentationSetting> chartSetting) : base(slide, chartSetting)
		{
			errorMessage = new TextBox(new TextBoxSetting()
			{
				textBlocks = new List<TextBlock>() { new TextBlock() { textValue = "This chart is not supported in this version of PowerPoint. Requires PowerPoint 2016 or later.", } }.ToArray(),
				x = chartSetting.applicationSpecificSetting.X,
				y = chartSetting.applicationSpecificSetting.Y,
				width = chartSetting.applicationSpecificSetting.Width,
				height = chartSetting.applicationSpecificSetting.Height,
			});
		}
		/// <summary>
		///
		/// </summary>
		internal void CreateExtendedChartGraphicFrame(string relationshipId, uint id, HyperlinkProperties hyperlinkProperties)
		{
			// Load Chart Part To Graphics Frame For Export
			P.NonVisualGraphicFrameProperties nonVisualProperties = new P.NonVisualGraphicFrameProperties()
			{
				NonVisualDrawingProperties = new P.NonVisualDrawingProperties { Id = id, Name = string.Format("Chart {0}", id) },
				NonVisualGraphicFrameDrawingProperties = new P.NonVisualGraphicFrameDrawingProperties(),
				ApplicationNonVisualDrawingProperties = new P.ApplicationNonVisualDrawingProperties()
			};
			if (hyperlinkProperties != null)
			{
				nonVisualProperties.NonVisualDrawingProperties.InsertAt(CreateHyperLink(hyperlinkProperties), 0);
			}
			graphicFrame = new P.GraphicFrame()
			{
				NonVisualGraphicFrameProperties = nonVisualProperties,
				Transform = new P.Transform(
				   new A.Offset
				   {
					   X = chartSetting.applicationSpecificSetting.X,
					   Y = chartSetting.applicationSpecificSetting.Y
				   },
				   new A.Extents
				   {
					   Cx = chartSetting.applicationSpecificSetting.Width,
					   Cy = chartSetting.applicationSpecificSetting.Height
				   }),
				Graphic = new A.Graphic(
				   new A.GraphicData(
					new CX.RelId() { Id = relationshipId }
				   )
				   {
					   Uri = "http://schemas.microsoft.com/office/drawing/2014/chartex"
				   })
			};
			CreateAlternateContent();
		}
		/// <summary>
		///
		/// </summary>
		public override void UpdateSize(uint width, uint height)
		{
			base.UpdateSize(width, height);
			errorMessage.UpdateSize(width, height);
		}
		/// <summary>
		///
		/// </summary>
		public override void UpdatePosition(uint x, uint y)
		{
			base.UpdatePosition(x, y);
			errorMessage.UpdatePosition(x, y);
		}
		private void CreateAlternateContent()
		{
			alternateContent = new AlternateContent(
				new AlternateContentChoice(
					(OpenXmlElement)graphicFrame.Clone()
				)
				{ Requires = "cx1" },
				new AlternateContentFallback(
					errorMessage.GetTextBoxShape()
				)
			);
			alternateContent.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
		}
		internal AlternateContent GetAlternateContent()
		{
			return alternateContent;
		}
	}
}
