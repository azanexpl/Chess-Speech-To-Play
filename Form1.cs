using System;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Speech.Recognition;
using System.Threading;
using System.Windows.Forms;

namespace ChessHelper
{
	public partial class Form1 : Form
	{
		private const int MouseeventfLeftdown = 0x02;
		private const int MouseeventfLeftup = 0x04;

		private readonly string[] _cyfry = new string[]
			{"one", "two", "three", "four", "five", "six", "seven", "eight"};

		private readonly int[,] _koordynatyPol = new int[64, 2];
		private readonly string[] _nazwyPol = new string[64];
		private const int Offset = 90;

		private readonly SpeechRecognitionEngine _recEngine = new SpeechRecognitionEngine();


		public Form1()
		{
			InitializeComponent();

			int licz = 0;
			int xBasic = 640;
			int yBasic = 850;


			for (var i = 0; i < 8; i++)
			for (var j = 0; j < 8; j++)
			{
				_nazwyPol[licz] = _cyfry[i] + " " + _cyfry[j];
				_koordynatyPol[licz, 0] = xBasic + j * Offset;
				_koordynatyPol[licz, 1] = yBasic - i * Offset;
				licz++;
			}

			licz = 0;
			for (var i = 0; i < 8; i++)
			{
				for (var j = 0; j < 8; j++)
				{
					Console.Write($@"({_koordynatyPol[licz, 0]}, {_koordynatyPol[licz, 1]}) ");
					licz++;
				}
				Console.WriteLine("");
			}

			
			var komendy = new Choices();
			komendy.Add(_nazwyPol);
			komendy.Add("close application", "resign", "start 10 minute game", "play", "white kingside castle",
				"white queenside castle", "black kingside castle", "black queenside castle");
			var slownik = new GrammarBuilder();
			slownik.Append(komendy);
			var gramatyka = new Grammar(slownik);

			_recEngine.LoadGrammar(gramatyka);
			_recEngine.SetInputToDefaultAudioDevice();
			_recEngine.SpeechRecognized += recEngine_SpeechRecognized;
			_recEngine.RecognizeAsync(RecognizeMode.Multiple);
		}

		[DllImport("user32.dll")]
		private static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);


		private void richTextBox1_TextChanged(object sender, EventArgs e)
		{
		}

		private void recEngine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
		{
			for (var i = 0; i < 64; i++)
				if (e.Result.Text == _nazwyPol[i])
				{
					Cursor.Position = new Point(_koordynatyPol[i, 0], _koordynatyPol[i, 1]);
					mouse_event(MouseeventfLeftdown, Cursor.Position.X, Cursor.Position.Y, 0, 0);
					mouse_event(MouseeventfLeftup, Cursor.Position.X, Cursor.Position.Y, 0, 0);
					richTextBox1.Text += $@"{_nazwyPol[i]} ({Cursor.Position.X}, {Cursor.Position.Y})";
					break;
				}

			switch (e.Result.Text)
			{
				case "white kingside castle":
					Cursor.Position = new Point(1000, 850);
					mouse_event(MouseeventfLeftdown, Cursor.Position.X, Cursor.Position.Y, 0, 0);
					mouse_event(MouseeventfLeftup, Cursor.Position.X, Cursor.Position.Y, 0, 0);
					Cursor.Position = new Point(1000 + 2 * Offset, 850);
					mouse_event(MouseeventfLeftdown, Cursor.Position.X, Cursor.Position.Y, 0, 0);
					mouse_event(MouseeventfLeftup, Cursor.Position.X, Cursor.Position.Y, 0, 0);
					break;
				case "white queenside castle":
					Cursor.Position = new Point(1000, 850);
					mouse_event(MouseeventfLeftdown, Cursor.Position.X, Cursor.Position.Y, 0, 0);
					mouse_event(MouseeventfLeftup, Cursor.Position.X, Cursor.Position.Y, 0, 0);
					Cursor.Position = new Point(1000 - 2 * Offset, 850);
					mouse_event(MouseeventfLeftdown, Cursor.Position.X, Cursor.Position.Y, 0, 0);
					mouse_event(MouseeventfLeftup, Cursor.Position.X, Cursor.Position.Y, 0, 0);
					break;
				case "black kingside castle":
					Cursor.Position = new Point(910, 850);
					mouse_event(MouseeventfLeftdown, Cursor.Position.X, Cursor.Position.Y, 0, 0);
					mouse_event(MouseeventfLeftup, Cursor.Position.X, Cursor.Position.Y, 0, 0);
					Cursor.Position = new Point(910 - 2 * Offset, 850);
					mouse_event(MouseeventfLeftdown, Cursor.Position.X, Cursor.Position.Y, 0, 0);
					mouse_event(MouseeventfLeftup, Cursor.Position.X, Cursor.Position.Y, 0, 0);
					break;
				case "black queenside castle":
					Cursor.Position = new Point(910, 850);
					mouse_event(MouseeventfLeftdown, Cursor.Position.X, Cursor.Position.Y, 0, 0);
					mouse_event(MouseeventfLeftup, Cursor.Position.X, Cursor.Position.Y, 0, 0);
					Cursor.Position = new Point(910 + 2 * Offset, 850);
					mouse_event(MouseeventfLeftdown, Cursor.Position.X, Cursor.Position.Y, 0, 0);
					mouse_event(MouseeventfLeftup, Cursor.Position.X, Cursor.Position.Y, 0, 0);
					break;
				case "resign":
					Cursor.Position = new Point(1600, 600);
					mouse_event(MouseeventfLeftdown, Cursor.Position.X, Cursor.Position.Y, 0, 0);
					mouse_event(MouseeventfLeftup, Cursor.Position.X, Cursor.Position.Y, 0, 0);
					mouse_event(MouseeventfLeftdown, Cursor.Position.X, Cursor.Position.Y, 0, 0);
					mouse_event(MouseeventfLeftup, Cursor.Position.X, Cursor.Position.Y, 0, 0);
					break;
				case "close application":
					Close();
					break;
				case "start 10 minute game":
					Process.Start("https://lichess.org");
					Thread.Sleep(2000);
					Cursor.Position = new Point(740, 176);
					mouse_event(MouseeventfLeftdown, Cursor.Position.X, Cursor.Position.Y, 0, 0);
					mouse_event(MouseeventfLeftup, Cursor.Position.X, Cursor.Position.Y, 0, 0);
					Cursor.Position = new Point(750, 660);
					mouse_event(MouseeventfLeftdown, Cursor.Position.X, Cursor.Position.Y, 0, 0);
					mouse_event(MouseeventfLeftup, Cursor.Position.X, Cursor.Position.Y, 0, 0);
					break;
			}
		}
	}
}