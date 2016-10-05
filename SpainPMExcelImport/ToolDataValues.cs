/*
 * Created by SharpDevelop.
 * User: val01039
 * Date: 5.10.2016
 * Time: 7:34
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;

namespace SpainPMExcelImport
{
	/// <summary>
	/// Description of ToolDataValues.
	/// </summary>
	public class ToolDataVAlues
	{
		public string _type;
		public string _unit;
		public string _description;
		public string _comment;
		public string _manufacturer;
		public string _product_id;	
		public string _number;	
		public string _diameter_offset;	
		public string _length_offset;	
		public string _break_control;
		public string _manual_tool_change;	
		public string _diameter;	
		public string _tip_diameter;	
		public string _tip_length;	
		public string _corner_radius;	
		public string _taper_angle;	
		public string _taper_angle2;	
		public string _flute_length;	
		public string _shoulder_length;	
		public string _shaft_diameter;	
		public string _body_length;	
		public string _overall_length;	
		public string _number_of_flutes;	
		public string _thread_pitch;	
		public string _coolant_support;	
		public string _coolant_mode;	
		public string _material_name;	
		public string _spindle_rpm;	
		public string _ramp_spindle_rpm;	
		public string _clockwise;	
		public string _cutting_feedrate;	
		public string _entry_feedrate;	
		public string _exit_feedrate;	
		public string _plunge_feedrate;	
		public string _ramp_feedrate;	
		public string _retract_feedrate;	
		public string _holder;	
		public string _shaft;	
		public string _guid;	
		public string _holder_description;	
		public string _holder_comment;	
		public string _holder_vendor;	
		public string _holder_product_id;	
		public string _holder_guid;	
		public string _holder_library_name;
		
		public ToolDataVAlues(string[] inputData)
		{
			_type = inputData[0];
			 _unit = inputData[1];
			 _description = inputData[2];
			 _comment = inputData[3];
			 _manufacturer = inputData[4];
			 _product_id = inputData[5];	
			 _number = inputData[6];	
			 _diameter_offset = inputData[7];	
			 _length_offset = inputData[8];	
			 _break_control = inputData[9];
			 _manual_tool_change = inputData[10];	
			 _diameter = inputData[11];	
			 _tip_diameter = inputData[12];	
			 _tip_length = inputData[13];	
			 _corner_radius = inputData[14];	
			 _taper_angle = inputData[15];	
			 _taper_angle2 = inputData[16];	
			 _flute_length = inputData[17];	
			 _shoulder_length = inputData[18];	
			 _shaft_diameter = inputData[19];	
			 _body_length = inputData[20];	
			 _overall_length = inputData[21];	
			 _number_of_flutes = inputData[22];	
			 _thread_pitch = inputData[23];	
			 _coolant_support = inputData[24];	
			 _coolant_mode = inputData[25];	
			 _material_name = inputData[26];	
			 _spindle_rpm = inputData[27];	
			 _ramp_spindle_rpm = inputData[28];	
			 _clockwise = inputData[29];	
			 _cutting_feedrate = inputData[30];	
			 _entry_feedrate = inputData[31];	
			 _exit_feedrate = inputData[32];	
			 _plunge_feedrate = inputData[33];	
			 _ramp_feedrate = inputData[34];	
			 _retract_feedrate = inputData[35];	
			 _holder = inputData[36];	
			 _shaft = inputData[37];	
			 _guid = inputData[38];	
			 _holder_description = inputData[39];	
			 _holder_comment = inputData[40];	
			 _holder_vendor = inputData[41];	
			 _holder_product_id = inputData[42];	
			 _holder_guid = inputData[43];	
			 _holder_library_name = inputData[44];
		}
	}
}
