// Silent, isolated WinForms + Core Audio fixture. Never controls other apps.
using System;
using System.IO;
using System.Media;
using System.Windows.Forms;

class AudioFixture : Form
{
    SoundPlayer player;
    Timer timer = new Timer();
    string stopFile;
    AudioFixture(string name, string stop)
    {
        Text = name;
        Width = 460;
        Height = 260;
        stopFile = stop;
        Controls.Add(new Label { Text = name + "\nSilent Core Audio test session", Dock = DockStyle.Fill,
                                TextAlign = System.Drawing.ContentAlignment.MiddleCenter });
        var memory = new MemoryStream();
        var writer = new BinaryWriter(memory);
        int bytes = 44100 * 2;
        writer.Write(System.Text.Encoding.ASCII.GetBytes("RIFF"));
        writer.Write(36 + bytes);
        writer.Write(System.Text.Encoding.ASCII.GetBytes("WAVEfmt "));
        writer.Write(16); writer.Write((short)1); writer.Write((short)1);
        writer.Write(44100); writer.Write(88200); writer.Write((short)2); writer.Write((short)16);
        writer.Write(System.Text.Encoding.ASCII.GetBytes("data")); writer.Write(bytes);
        writer.Write(new byte[bytes]); memory.Position = 0;
        player = new SoundPlayer(memory);
        player.PlayLooping();
        timer.Interval = 100;
        timer.Tick += (sender, args) => { if (File.Exists(stopFile)) Close(); };
        timer.Start();
        FormClosed += (sender, args) => { player.Stop(); player.Dispose(); timer.Dispose(); };
    }
    [STAThread]
    static void Main(string[] args)
    {
        Application.EnableVisualStyles();
        Application.Run(new AudioFixture(args[0], args[1]));
    }
}
