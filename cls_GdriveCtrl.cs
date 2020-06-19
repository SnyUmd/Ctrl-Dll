using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;

namespace Ctrl_Dll
{
    class cls_GdriveCtrl
    {
        //******************************************************
        public void mGD_Download()
        {
            var fileId = "0BwwA4oUTeiV1UVNwOHItT0xfa2M";
            var request = driveService.Files.Get(fileId);
            var stream = new System.IO.MemoryStream();

            
        }
    }
}
