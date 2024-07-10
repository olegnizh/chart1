/*
 * Created by SharpDevelop.
 * User: alexpc
 * Date: 24.02.2023
 * Time: 14:02
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Data;
using System.Data.OleDb;

namespace chart1
{
	/// <summary>
	/// Description of Info.
	/// </summary>
	public static class Info
	{
        public static string Name { get; set; }
        public static string LastName { get; set; }
        //public static string SecName { get; set; }

        public static string path_app = "";
        public static string path_data = "";

        public static DataTable dt = new DataTable();		
		
        // config
        // browser
        public static string s_browser = "";
        // work folder
        public static string s_work = "";
        public static string s_name = "";
        public static string s_work_a = "";
		
	}
}
