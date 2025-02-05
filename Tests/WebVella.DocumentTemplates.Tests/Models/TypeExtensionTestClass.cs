namespace WebVella.DocumentTemplates.Tests.Models;
public class WvTestsSimpleClass { }
public class WvTestsSimpleGenericClass<T> { }
public interface IWvTestsSimpleInterface { }
public interface IWvTestsGenericInterface<T> { }
public class WvTestsSimpleBaseClass { }
public class WvTestsGenericBaseClass<T> { }
public class WvTestsChildFromSimpleBaseClass : WvTestsSimpleBaseClass { }
public class WvTestsChildFromGenericBaseClass : WvTestsGenericBaseClass<WvTestsSimpleClass> { }
public class WVTestsClassImplementsSimpleInterfaceClass : IWvTestsSimpleInterface { }
public class WVTestsClassImplementsGenericInterfaceClass : IWvTestsGenericInterface<WvTestsSimpleClass> { }
