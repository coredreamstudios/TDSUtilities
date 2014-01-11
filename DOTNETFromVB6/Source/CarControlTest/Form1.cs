using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace CarControlTest
{
	public partial class Form1 : Form
	{
		Random random = new Random();
		public Form1()
		{
			this.InitializeComponent();
		}

		private void timer_Tick(object sender, EventArgs e)
		{
			this.car.FrontL = random.NextDouble();
			this.car.FrontR = random.NextDouble();
			this.car.RearL = random.NextDouble();
			this.car.RearR = random.NextDouble();
		}
	}
}