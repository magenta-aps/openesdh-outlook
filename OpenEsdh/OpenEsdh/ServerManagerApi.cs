namespace OpenEsdh
{
    using SharpShell;
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.Composition.Hosting;
    using System.IO;
    using System.Windows.Forms;

    public static class ServerManagerApi
    {
        public static IEnumerable<ServerEntry> LoadServers(string path)
        {
            List<ServerEntry> list = new List<ServerEntry>();
            try
            {
                AssemblyCatalog catalog = new AssemblyCatalog(Path.GetFullPath(path));
                IEnumerable<Lazy<ISharpShellServer>> exports = new CompositionContainer(catalog, new ExportProvider[0]).GetExports<ISharpShellServer>();
                foreach (Lazy<ISharpShellServer> lazy in exports)
                {
                    ISharpShellServer server = null;
                    try
                    {
                        server = lazy.Value;
                    }
                    catch (Exception)
                    {
                        ServerEntry entry = new ServerEntry {
                            ServerName = "Invalid",
                            ServerPath = path,
                            ServerType = ServerType.None
                        };
                        Guid guid = new Guid();
                        entry.ClassId = guid;
                        entry.Server = null;
                        entry.IsInvalid = true;
                        list.Add(entry);
                        continue;
                    }
                    ServerEntry item = new ServerEntry {
                        ServerName = server.DisplayName,
                        ServerPath = path,
                        ServerType = server.ServerType,
                        ClassId = server.ServerClsid,
                        Server = server
                    };
                    list.Add(item);
                }
            }
            catch (Exception)
            {
                MessageBox.Show("The file '" + Path.GetFileName(path) + "' is not a SharpShell Server.", "Warning");
            }
            return list;
        }
    }
}

