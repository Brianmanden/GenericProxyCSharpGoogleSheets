using System;
using System.Reflection;

namespace GenericProxyCSharpGoogleSheets.Areas.HelpPage.ModelDescriptions
{
    public interface IModelDocumentationProvider
    {
        string GetDocumentation(MemberInfo member);

        string GetDocumentation(Type type);
    }
}