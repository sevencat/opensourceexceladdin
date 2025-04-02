using Nancy;

namespace Sevencat.AddinFunc.Ui.Service;

public class TestService : NancyModule
{
	public TestService() : base("/api/test")
	{
		Get("/hello", _ => "Hello from Nancy");
	}
}