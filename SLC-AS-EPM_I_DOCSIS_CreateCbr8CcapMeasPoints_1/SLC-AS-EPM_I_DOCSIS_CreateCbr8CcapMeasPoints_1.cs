// Ignore Spelling: Cbr Ccap

namespace SLCASEPMIDOCSISCreateCbr8CcapMeasPoints_1
{
	using System;
	using System.Collections.Generic;
	using Skyline.DataMiner.Automation;
	using Skyline.DataMiner.Core.DataMinerSystem.Automation;
	using Skyline.DataMiner.Core.DataMinerSystem.Common;

	/// <summary>
	/// Represents a DataMiner Automation script.
	/// </summary>
	public class Script
	{
		/// <summary>
		/// The script entry point.
		/// </summary>
		/// <param name="engine">Link with SLAutomation process.</param>
		public static void Run(IEngine engine)
		{
			try
			{
				RunSafe(engine);
			}
			catch (ScriptAbortException)
			{
				// Catch normal abort exceptions (engine.ExitFail or engine.ExitSuccess)
				throw; // Comment if it should be treated as a normal exit of the script.
			}
			catch (ScriptForceAbortException)
			{
				// Catch forced abort exceptions, caused via external maintenance messages.
				throw;
			}
			catch (ScriptTimeoutException)
			{
				// Catch timeout exceptions for when a script has been running for too long.
				throw;
			}
			catch (InteractiveUserDetachedException)
			{
				// Catch a user detaching from the interactive script by closing the window.
				// Only applicable for interactive scripts, can be removed for non-interactive scripts.
				throw;
			}
			catch (Exception e)
			{
				engine.ExitFail("Run|Something went wrong: " + e);
			}
		}

		private static void RunSafe(IEngine engine)
		{
			var dms = engine.GetDms();
			var elements = dms.GetElements();
			foreach (var element in elements)
			{
				if (!element.Protocol.Name.Equals("CISCO CBR-8 CCAP UTSC") || !element.Protocol.Version.Equals("Production"))
				{
					continue;
				}

				var ports = element.GetStandaloneParameter<string>(3001).GetValue().Split(';');
				var itemId = 1;
				var dmaId = Convert.ToString(element.AgentId);
				var elementId = Convert.ToString(element.Id);
				var elementName = element.Name;

				var newMeasPoints = new List<object>();

				foreach (var port in ports)
				{
					var itemStringId = Convert.ToString(itemId);
					string[] measPoint = new[]
					{
						itemStringId, // Measurement Point ID (Unique)
						string.Format("{0};{0};{0};{0};{0}", dmaId), // DMA id of parameter to set
						string.Format("{0};{0};{0};{0};{0}", elementId), // Element id(s) of parameter(s) to set
						"3100;3108;3109;3124;3325", // Write parameter id(s) to set
						$"{elementName}_{port}", // Measurement point name
						"true;false;false;true;false", // For each parameter to set, "true" if it is of string type
						$"{port};35;60;1;1", // Value(s) to set in parameter(s)
						"0", // Delay time after set
						"3000;3008;3009;3024;3225", // Read id(s) associated with the write parameter ids specified above
						";;;;", // Table indexes of the parameters to set
						"0.000000", // Freq offset
						"false", // Needs invert freq
						string.Empty, // Automation script info (to be used instead of parameter sets)
						string.Empty, // Amplitude correction (RN3223)
					};

					itemId++;

					newMeasPoints.Add(measPoint);
				}

				element.SpectrumAnalyzer.SetMeasurementPoints(false, newMeasPoints.ToArray());
			}
		}
	}
}
