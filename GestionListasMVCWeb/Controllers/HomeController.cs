using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Security;
using System.Web;
using System.Web.Mvc;
using GestionListasMVCWeb.Models;

namespace GestionListasMVCWeb.Controllers
{
    public class HomeController : Controller
    {
        [SharePointContextFilter]
        public ActionResult Index()
        {

            Session["SPHostUrl"] = Request.QueryString["SPHostUrl"];
            Session["SPLanguage"] = Request.QueryString["SPLanguage"];
            Session["SPClientTag"] = Request.QueryString["SPClientTag"];
            Session["SPProductNumber"] = Request.QueryString["SPProductNumber"];

            User spUser = null;

            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {
                    spUser = clientContext.Web.CurrentUser;

                    clientContext.Load(spUser, user => user.Title);

                    clientContext.ExecuteQuery();

                    ViewBag.UserName = spUser.Title;
                }
            }

            return View();
        }

        public ActionResult Alta()
        {
            return View(new Alumno());
        }

        [HttpPost]
        public ActionResult Alta(Alumno alumno)
        {
           
            using (var context = new ClientContext("https://letias.sharepoint.com"))
            {
                var pwd = new SecureString();
                foreach (var c in "Cursill0".ToCharArray())
                {
                    pwd.AppendChar(c);
                }

                context.Credentials = new SharePointOnlineCredentials("‎learrsan@letias.onmicrosoft.com", pwd);
               
                List list = context.Web.Lists.GetByTitle("Alumnos");
                context.Load(list);
                var creacion = new ListItemCreationInformation();
                var item = list.AddItem(creacion);
                item["Title"] = alumno.Nombre;
                item["Apellido"] = alumno.Apellido;
                item["Edad"] = alumno.Edad;
                item["Nota"] = alumno.Nota;

                item.Update();

                context.ExecuteQuery();
            }
            String redireccion = String.Format("/?SPHostUrl={0}&SPLanguage={1}&SPClientTag{2}&" +
                                               "SPProductNumber={3}", Session["SPHostUrl"],
                Session["SPLanguage"], Session["SPClientTag"],
                Session["SPProductNumber"]);

            return RedirectToAction("Index","Home");
        }
    }
}
