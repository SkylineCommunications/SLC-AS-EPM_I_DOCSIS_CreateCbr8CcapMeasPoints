// Ignore Spelling: Cbr Ccap

namespace SLCASEPMIDOCSISCreateCbr8CcapMeasPoints_1
{
	using System;
	using System.Collections.Generic;
	using System.Linq;
	using Skyline.DataMiner.Automation;
	using Skyline.DataMiner.Core.DataMinerSystem.Automation;
	using Skyline.DataMiner.Core.DataMinerSystem.Common;
	using Skyline.DataMiner.Net.Helper;

	/// <summary>
	/// Represents a DataMiner Automation script.
	/// </summary>
	public class Script
	{
		public enum Cbr8CcapParams
		{
			InterfacesList = 3001,

			UpPortRead = 3000,
			UpPortWrite = 3100,

			CenterFreqRead = 3008,
			CenterFreqWrite = 3108,

			FreqSpanRead = 3009,
			FreqSpanWrite = 3109,

			DestIdxRead = 3024,
			DestIdxWrite = 3124,

			InitCaptureRead = 3225,
			InitCaptureWrite = 3325,
		}

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
			var elements = dms.GetElements().ToList();
			foreach (var element in elements.Where(x => x.Name.Equals("CISCO CBR-8 CCAP UTSC") && x.Protocol.Version.Equals("Production")))
			{
				var portsString = element.GetStandaloneParameter<string>((int)Cbr8CcapParams.InterfacesList).GetValue();
				if (portsString.IsNullOrEmpty())
                {
					continue;
				}

				var ports = portsString.Split(';');
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
						$"{(int)Cbr8CcapParams.UpPortWrite};{(int)Cbr8CcapParams.CenterFreqWrite};{(int)Cbr8CcapParams.FreqSpanWrite};{(int)Cbr8CcapParams.DestIdxWrite};{(int)Cbr8CcapParams.InitCaptureWrite}", // Write parameter id(s) to set
						$"{elementName}_{port}", // Measurement point name
						"true;false;false;true;false", // For each parameter to set, "true" if it is of string type
						$"{port};35;60;1;1", // Value(s) to set in parameter(s)
						"0", // Delay time after set
						$"{(int)Cbr8CcapParams.UpPortRead};{(int)Cbr8CcapParams.CenterFreqRead};{(int)Cbr8CcapParams.FreqSpanRead};{(int)Cbr8CcapParams.DestIdxRead};{(int)Cbr8CcapParams.InitCaptureRead}", // Read id(s) associated with the write parameter ids specified above
						";;;;", // Table indexes of the parameters to set
						"0.000000", // Freq offset
						"false", // Needs invert freq
						string.Empty, // Automation script info (to be used instead of parameter sets)
						string.Empty, // Amplitude correction (RN3223)
					};

					itemId++;

					newMeasPoints.Add(measPoint);
				}

				try
				{
					element.SpectrumAnalyzer.SetMeasurementPoints(false, newMeasPoints.ToArray());
				}
				catch (Exception e)
				{
					engine.Log($"An error has occurred while setting Measurement Points in the {element.Name} element. Error: {e}");
				}
			}
		}
	}
}
