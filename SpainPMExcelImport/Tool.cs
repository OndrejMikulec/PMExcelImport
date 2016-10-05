/*
 * Created by SharpDevelop.
 * User: val01039
 * Date: 5.10.2016
 * Time: 7:37
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;

namespace SpainPMExcelImport
{
	/// <summary>
	/// Description of Tool.
	/// </summary>
	public class Tool
	{
		PowerMILL.Application _pmApp;
		
		string _name;
		string _id;
		bool _numberUserDefined;
		int _numberValue;
		new enum Type {end_mill,tip_radiused,ball_nosed,taper_spherical,taper_tipped,off_centre_tip_rad,tipped_disc,drill,tap,form,routing,thread_mill,barrel,dovetail}
		Type _type;
		double _length;
		string _identifier;
		double _diameter;
		double _upperTipRadius;
		double _barrelRadius;
		bool _flatBottom;
		string _status;
		double _overhang;
		double _pitch;
		double _tipRadius;
		double _tipRadiusCentre;
		double _taperAngle;
		double _taperHeight;
		int _numberOfFlutes;
		string _description;
		double _routinEndMillDiameter;
		
		double _axialDepthOfCutFinishingCopy_milling;
		double _axialDepthOfCutFinishingDrill;
		double _axialDepthOfCutFinishingFace_milling;
		double _axialDepthOfCutFinishingGeneral;
		double _axialDepthOfCutFinishingPlunge_milling;
		double _axialDepthOfCutFinishingProfiling;
		double _axialDepthOfCutFinishingSlotting;
		
		double _axialDepthOfCutRoughingCopy_milling;
		double _axialDepthOfCutRoughingDrill;
		double _axialDepthOfCutRoughingFace_milling;
		double _axialDepthOfCutRoughingGeneral;
		double _axialDepthOfCutRoughingPlunge_milling;
		double _axialDepthOfCutRoughingProfiling;
		double _axialDepthOfCutRoughingSlotting ;
		
		double _radialDepthOfCutFinishingCopy_milling;
		double _radialDepthOfCutFinishingDrill;
		double _radialDepthOfCutFinishingFace_milling;
		double _radialDepthOfCutFinishingGeneral;
		double _radialDepthOfCutFinishingPlunge_milling;
		double _radialDepthOfCutFinishingProfiling;
		double _radialDepthOfCutFinishingSlotting;
		
		double _radialDepthOfCutRoughingCopy_milling;
		double _radialDepthOfCutRoughingDrill;
		double _radialDepthOfCutRoughingFace_milling;
		double _radialDepthOfCutRoughingGeneral;
		double _radialDepthOfCutRoughingPlunge_milling;
		double _radialDepthOfCutRoughingProfiling;
		double _radialDepthOfCutRoughingSlotting ;
		
		double _feedPerToothFinishingCopy_milling;
		double _feedPerToothFinishingDrill;
		double _feedPerToothFinishingFace_milling;
		double _feedPerToothFinishingGeneral;
		double _feedPerToothFinishingPlunge_milling;
		double _feedPerToothFinishingProfiling;
		double _feedPerToothFinishingSlotting;
		
		double _feedPerToothRoughingCopy_milling;
		double _feedPerToothRoughingDrill;
		double _feedPerToothRoughingFace_milling;
		double _feedPerToothRoughingGeneral;
		double _feedPerToothRoughingPlunge_milling;
		double _feedPerToothRoughingProfiling;
		double _feedPerToothRoughingSlotting ;
		
		double _cuttingSpeedFinishingCopy_milling;
		double _cuttingSpeedFinishingDrill;
		double _cuttingSpeedFinishingFace_milling;
		double _cuttingSpeedFinishingGeneral;
		double _cuttingSpeedFinishingPlunge_milling;
		double _cuttingSpeedFinishingProfiling;
		double _cuttingSpeedFinishingSlotting;
		
		double _cuttingSpeedRoughingCopy_milling;
		double _cuttingSpeedRoughingDrill;
		double _cuttingSpeedRoughingFace_milling;
		double _cuttingSpeedRoughingGeneral;
		double _cuttingSpeedRoughingPlunge_milling;
		double _cuttingSpeedRoughingProfiling;
		double _cuttingSpeedRoughingSlotting ;
		
		string _toolFamily;
		new enum Coolant {none,standard,flood,mist,tap,air,thru,both}
		Coolant _coolant;
		string _stockMaterial;
		bool _idTracksName;
		string _holderName;
		
		struct Segment
		{
			public double UpperDiameter {get;set;}
			public double LowerDiameter{get;set;}
			public double Length{get;set;}
			public bool Ignore{get;set;}
		}
		
		List<Segment> _shankSegments = new List<Segment>();
		List<Segment> _holderSegments = new List<Segment>();
		
		
		public Tool(ToolDataVAlues inputData, PowerMILL.Application pmApp)
		{
			_pmApp = pmApp;
			
			FillFields(inputData);
			
			applyToPm();
			
		}

		void FillFields(ToolDataVAlues inputData)
		{
			_name = inputData._type+inputData._diameter + inputData._number;
			_id = _name;
			_numberUserDefined = true;
			_numberValue = int.Parse( inputData._number);
			
			switch (inputData._type) {
				case "ball end mill" :
					_type = Type.ball_nosed;
					break;
				case "bull nose end mill" :
					_type = Type.tip_radiused;
					break;
				case "face mill" :
					_type = Type.tip_radiused;
					break;
			}
			

			_length = double.Parse( inputData._flute_length);
			_identifier =_name;
			_diameter = double.Parse( inputData._diameter);
			_upperTipRadius;
			_barrelRadius;
			_flatBottom;
			_status = "";
			_overhang = double.Parse( inputData._shoulder_length);
			_pitch;
			_tipRadius = double.Parse( inputData._corner_radius);
			_tipRadiusCentre;
			_taperAngle;
			_taperHeight;
			_numberOfFlutes = int.Parse( inputData._number_of_flutes);
			_description;
			_routinEndMillDiameter;
			
			
			
			_axialDepthOfCutFinishingCopy_milling = 0;
			 _axialDepthOfCutFinishingDrill = 0;
			 _axialDepthOfCutFinishingFace_milling = 0;
			 _axialDepthOfCutFinishingGeneral = 0;
			 _axialDepthOfCutFinishingPlunge_milling = 0;
			 _axialDepthOfCutFinishingProfiling  = 0;
			 _axialDepthOfCutFinishingSlotting  = 0;
			
			 _axialDepthOfCutRoughingCopy_milling = 0;
			 _axialDepthOfCutRoughingDrill = 0;
			 _axialDepthOfCutRoughingFace_milling = 0;
			 _axialDepthOfCutRoughingGeneral = 0;
			 _axialDepthOfCutRoughingPlunge_milling = 0;
			 _axialDepthOfCutRoughingProfiling = 0;
			 _axialDepthOfCutRoughingSlotting  = 0;
			
			 _radialDepthOfCutFinishingCopy_milling = 0;
			 _radialDepthOfCutFinishingDrill = 0;
			 _radialDepthOfCutFinishingFace_milling = 0;
			 _radialDepthOfCutFinishingGeneral = 0;
			 _radialDepthOfCutFinishingPlunge_milling = 0;
			 _radialDepthOfCutFinishingProfiling = 0;
			 _radialDepthOfCutFinishingSlotting = 0;
			
			 _radialDepthOfCutRoughingCopy_milling = 0;
			 _radialDepthOfCutRoughingDrill = 0;
			 _radialDepthOfCutRoughingFace_milling = 0;
			 _radialDepthOfCutRoughingGeneral = 0;
			 _radialDepthOfCutRoughingPlunge_milling = 0;
			 _radialDepthOfCutRoughingProfiling = 0;
			 _radialDepthOfCutRoughingSlotting  = 0;
			 
			 double ft = double.Parse(inputData._cutting_feedrate)/(double.Parse(inputData._spindle_rpm)*int.Parse( inputData._number_of_flutes));
			 double ftPlunge = double.Parse(inputData._plunge_feedrate)/(double.Parse(inputData._ramp_spindle_rpm)*int.Parse( inputData._number_of_flutes));
			
			 _feedPerToothFinishingCopy_milling = ft;
			 _feedPerToothFinishingDrill = ft;
			 _feedPerToothFinishingFace_milling = ft;
			 _feedPerToothFinishingGeneral = ft;
			 _feedPerToothFinishingPlunge_milling = ftPlunge;
			 _feedPerToothFinishingProfiling = ft;
			 _feedPerToothFinishingSlotting = ft;
			
			 _feedPerToothRoughingCopy_milling = ft;
			 _feedPerToothRoughingDrill = ft;
			 _feedPerToothRoughingFace_milling = ft;
			 _feedPerToothRoughingGeneral = ft;
			 _feedPerToothRoughingPlunge_milling = ftPlunge;
			 _feedPerToothRoughingProfiling = ft;
			 _feedPerToothRoughingSlotting  = ft;
			 
			 double cs = (Math.PI*_diameter*double.Parse(inputData._spindle_rpm))/1000;
			 double csPlunge = (Math.PI*_diameter*double.Parse(inputData._ramp_spindle_rpm))/1000;
			
			 _cuttingSpeedFinishingCopy_milling = cs;
			 _cuttingSpeedFinishingDrill = cs;
			 _cuttingSpeedFinishingFace_milling = cs;
			 _cuttingSpeedFinishingGeneral = cs;
			 _cuttingSpeedFinishingPlunge_milling = csPlunge;
			 _cuttingSpeedFinishingProfiling = cs;
			 _cuttingSpeedFinishingSlotting = cs;
			
			 _cuttingSpeedRoughingCopy_milling = cs;
			 _cuttingSpeedRoughingDrill = cs;
			 _cuttingSpeedRoughingFace_milling = cs;
			 _cuttingSpeedRoughingGeneral = cs;
			 _cuttingSpeedRoughingPlunge_milling = csPlunge;
			 _cuttingSpeedRoughingProfiling = cs;
			 _cuttingSpeedRoughingSlotting  = cs;
			 
			 _toolFamily;
			
			
			switch (inputData._coolant_mode) {
				case "flood" :
					_coolant = Coolant.flood;
					break;
				
			}
			
			
			
			_stockMaterial;
			_idTracksName = true;
			_holderName;
			
			_shankSegments.Add(new Segment(){UpperDiameter = _diameter-0.1, LowerDiameter=_diameter-0.1,Length=double.Parse(inputData._shoulder_length),Ignore=false});
			_shankSegments.Add(new Segment(){UpperDiameter = double.Parse(inputData._shaft_diameter) -0.1, double.Parse(inputData._shaft_diameter) -0.1,Length=double.Parse(inputData._body_length)-double.Parse(inputData._shoulder_length),Ignore=false});
			
			 
		}
		
		void applyToPm()
		{
			_pmApp.DoCommand("CREATE TOOL ;");
			_pmApp.DoCommand(@"$entity(""Tool"","""").Diameter = "+_PrumerD1.ToString().Replace(",","."));
			_pmApp.DoCommand(@"$entity(""Tool"","""+_NazevN+@""").NumberOfFlutes = "+_Zuby);
			_pmApp.DoCommand(@"$entity(""Tool"","""+_NazevN+@""").Coolant = "+PMCoolant);
			_pmApp.DoCommand(@"$entity(""Tool"","""+_NazevN+@""").Number.UserDefined = 1");
			_pmApp.DoCommand(@"$entity(""Tool"","""+_NazevN+@""").Number.Value = "+_Nastroj);
			_pmApp.DoCommand(@"$entity(""Tool"","""+_NazevN+@""").Overhang = "+_Vylozeni.ToString().Replace(",","."));
			_pmApp.DoCommand(@"$entity(""Tool"","""+_NazevN+@""").Type = ""drill""");
			_pmApp.DoCommand(@"$entity(""Tool"","""+_NazevN+@""").TipRadius = "+_RadiusRohu.ToString().Replace(",","."));
			_pmApp.DoCommand(@"$entity(""Tool"","""+_NazevN+@""").Length = "+Segmenty[0].Rozmer2.ToString().Replace(",","."));
			_pmApp.DoCommand(@"EDIT TOOL  """+_NazevN+@""" SHANK_CLEAR");
			
			for (int i = 0; i < Segmenty.Count; i++) {
				slozenaDelka += Segmenty[i].Rozmer2;
				if (slozenaDelka>Segmenty[i].Rozmer2) {
					
					
					_pmApp.DoCommand(@"EDIT TOOL  """+_NazevN+@"""  SHANK_COMPONENT ADD");
					_pmApp.DoCommand(@"EDIT TOOL  """+_NazevN+@"""  SHANK_COMPONENT UPPERDIA "+Segmenty[i].Rozmer1.ToString().Replace(".",","));
					_pmApp.DoCommand(@"EDIT TOOL  """+_NazevN+@"""  SHANK_COMPONENT LOWERDIA "+Segmenty[i].Rozmer3.ToString().Replace(".",","));
					_pmApp.DoCommand(@"EDIT TOOL  """+_NazevN+@"""  SHANK_COMPONENT LENGTH "+Segmenty[i].Rozmer2.ToString().Replace(".",","));
					
					
				}
				
				
			}
			
		}
	}
}
