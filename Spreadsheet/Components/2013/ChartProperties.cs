// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Global_2013;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Spreadsheet_2013
{
	/// <summary>
	///
	/// </summary>
	public class ChartProperties
	{
		/// <summary>
		///
		/// </summary>
		internal readonly ChartSetting chartSetting;
		/// <summary>
		///
		/// </summary>
		internal readonly Worksheet currentWorksheet;
		/// <summary>
		///
		/// </summary>
		internal P.GraphicFrame? graphicFrame;
		/// <summary>
		///
		/// </summary>
		internal ChartProperties(Worksheet worksheet, ChartSetting chartSetting)
		{
			this.chartSetting = chartSetting;
			currentWorksheet = worksheet;
		}
		
		/// <summary>
		/// </summary>
		/// <returns>
		/// X,Y
		/// </returns>
		internal (uint, uint) GetPosition()
		{
			return (chartSetting.x, chartSetting.y);
		}

		/// <summary>
		/// </summary>
		/// <returns>
		/// Width,Height
		/// </returns>
		internal (uint, uint) GetSize()
		{
			return (chartSetting.width, chartSetting.height);
		}

		/// <summary>
		/// Save Chart Part
		/// </summary>
		internal void Save()
		{
			currentWorksheet.GetWorksheetPart().Worksheet.Save();
		}

		/// <summary>
		/// Update Chart Position
		/// </summary>
		/// <param name="X">
		/// </param>
		/// <param name="Y">
		/// </param>
		public virtual void UpdatePosition(uint X, uint Y)
		{
			chartSetting.x = X;
			chartSetting.y = Y;
			if (graphicFrame != null)
			{
				graphicFrame.Transform = new P.Transform
				{
					Offset = new A.Offset { X = chartSetting.x, Y = chartSetting.y },
					Extents = new A.Extents { Cx = chartSetting.width, Cy = chartSetting.height }
				};
			}
		}

		/// <summary>
		/// Update Chart Size
		/// </summary>
		/// <param name="Width">
		/// </param>
		/// <param name="Height">
		/// </param>
		public virtual void UpdateSize(uint Width, uint Height)
		{
			chartSetting.width = Width;
			chartSetting.height = Height;
			if (graphicFrame != null)
			{
				graphicFrame.Transform = new P.Transform
				{
					Offset = new A.Offset { X = chartSetting.x, Y = chartSetting.y },
					Extents = new A.Extents { Cx = chartSetting.width, Cy = chartSetting.height }
				};
			}
		}

		internal void CreateChartGraphicFrame(string relationshipId, uint id)
		{
			// Load Chart Part To Graphics Frame For Export
			P.NonVisualGraphicFrameProperties nonVisualProperties = new()
			{
				NonVisualDrawingProperties = new P.NonVisualDrawingProperties { Id = id, Name = "Chart" },
				NonVisualGraphicFrameDrawingProperties = new P.NonVisualGraphicFrameDrawingProperties(),
				ApplicationNonVisualDrawingProperties = new P.ApplicationNonVisualDrawingProperties()
			};
			graphicFrame = new()
			{
				NonVisualGraphicFrameProperties = nonVisualProperties,
				Transform = new P.Transform(
				   new A.Offset
				   {
					   X = chartSetting.x,
					   Y = chartSetting.y
				   },
				   new A.Extents
				   {
					   Cx = chartSetting.width,
					   Cy = chartSetting.height
				   }),
				Graphic = new A.Graphic(
				   new A.GraphicData(
					   new C.ChartReference { Id = relationshipId }
				   )
				   { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" })
			};
		}

		internal P.GraphicFrame GetChartGraphicFrame()
		{
			return graphicFrame!;
		}

	}

}
