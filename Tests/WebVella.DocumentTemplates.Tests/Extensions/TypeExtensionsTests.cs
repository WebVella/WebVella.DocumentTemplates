using WebVella.DocumentTemplates.Extensions;
using WebVella.DocumentTemplates.Tests.Models;
namespace WebVella.DocumentTemplates.Tests.Extensions;
public class TypeExtensionsTests
{
	//Class extensions
	[Fact]
	public void Type_ShouldNotInheritClass()
	{
		Type objType = typeof(WvTestsSimpleClass);
		bool? result = null;
		var action = () => { result = objType.InheritsClass(typeof(WvTestsSimpleBaseClass)); };
		action();
		Assert.False(result);
	}

	[Fact]
	public void Type_ShouldInheritClass()
	{
		Type objType = typeof(WvTestsChildFromSimpleBaseClass);
		bool? result = null;
		var action = () => { result = objType.InheritsClass(typeof(WvTestsSimpleBaseClass)); };
		action();
		Assert.True(result);
	}

	[Fact]
	public void Type_ShouldNotInheritGenericClass()
	{
		Type objType = typeof(WvTestsChildFromGenericBaseClass);
		bool? result = null;
		var action = () => { result = objType.InheritsGenericClass(typeof(WvTestsGenericBaseClass<>), typeof(WvTestsSimpleBaseClass)); };
		action();
		Assert.False(result);
	}

	[Fact]
	public void Type_ShouldInheritGenericClass()
	{
		Type objType = typeof(WvTestsChildFromGenericBaseClass);
		bool? result = null;
		var action = () => { result = objType.InheritsGenericClass(typeof(WvTestsGenericBaseClass<>), typeof(WvTestsSimpleClass)); };
		action();
		Assert.True(result);
	}

	//Interface extensions
	[Fact]
	public void Type_ShouldNotImplementInterface()
	{
		Type objType = typeof(WvTestsSimpleClass);
		bool? result = null;
		var action = () => { result = objType.ImplementsInterface(typeof(IWvTestsSimpleInterface)); };
		action();
		Assert.False(result);
	}

	[Fact]
	public void Type_ShouldImplementInterface()
	{
		Type objType = typeof(WVTestsClassImplementsSimpleInterfaceClass);
		bool? result = null;
		var action = () => { result = objType.ImplementsInterface(typeof(IWvTestsSimpleInterface)); };
		action();
		Assert.True(result);
	}

	[Fact]
	public void Type_ShouldNotImplementGenericInterface()
	{
		Type objType = typeof(WVTestsClassImplementsSimpleInterfaceClass);
		bool? result = null;
		var action = () => { result = objType.ImplementsGenericInterface(typeof(IWvTestsGenericInterface<>), typeof(WvTestsSimpleClass)); };
		action();
		Assert.False(result);
	}

	[Fact]
	public void Type_ShouldImplementGenericInterface()
	{
		Type objType = typeof(WVTestsClassImplementsGenericInterfaceClass);
		bool? result = null;
		var action = () => { result = objType.ImplementsGenericInterface(typeof(IWvTestsGenericInterface<>), typeof(WvTestsSimpleClass)); };
		action();
		Assert.True(result);
	}

	[Fact]
	public void Type_GetGenericTypeFromGenericInterfaceFalse()
	{
		Type objType = typeof(WvTestsSimpleClass);
		string? result = null;
		var action = () => { result = objType.GetGenericTypeFullNameFromGenericInterface(); };
		var ex = Record.Exception(action);
		Assert.IsType<ArgumentException>(ex);
		Assert.Equal("The provided type must be a interface. (Parameter 'type')", ex.Message);
	}

	[Fact]
	public void Type_GetGenericTypeFromGenericInterfaceFalse2()
	{
		Type objType = typeof(IWvTestsSimpleInterface);
		string? result = null;
		var action = () => { result = objType.GetGenericTypeFullNameFromGenericInterface(); };
		var ex = Record.Exception(action);
		Assert.IsType<ArgumentException>(ex);
		Assert.Equal("The provided type must be a generic interface. (Parameter 'type')", ex.Message);
	}

	[Fact]
	public void Type_GetGenericTypeFromGenericInterfaceTrue()
	{
		Type objType = typeof(IWvTestsGenericInterface<WvTestsSimpleClass>);
		string? result = null;
		var action = () => { result = objType.GetGenericTypeFullNameFromGenericInterface(); };
		action();
		Assert.Equal(result, typeof(WvTestsSimpleClass).FullName);
	}

	[Fact]
	public void Type_GetGenericTypeFromImplementedGenericInterfaceTrue()
	{
			Type objType = typeof(WVTestsClassImplementsGenericInterfaceClass);
			List<string> result = new();
			var action = () => { result = objType.GetGenericTypeFullNameFromImplementedGenericInterface(typeof(IWvTestsGenericInterface<>)); };
			action();
			Assert.True(result.Count == 1);
			Assert.Equal(result[0],typeof(WvTestsSimpleClass).FullName);
	}
}