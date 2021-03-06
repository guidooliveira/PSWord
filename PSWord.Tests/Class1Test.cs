// <copyright file="Class1Test.cs">Copyright ©  2015</copyright>
using System;
using Microsoft.Pex.Framework;
using Microsoft.Pex.Framework.Validation;
using NUnit.Framework;
using PSWord;

namespace PSWord.Tests
{
    /// <summary>This class contains parameterized unit tests for Class1</summary>
    [PexClass(typeof(AddWordText))]
    [PexAllowedExceptionFromTypeUnderTest(typeof(InvalidOperationException))]
    [PexAllowedExceptionFromTypeUnderTest(typeof(ArgumentException), AcceptExceptionSubtypes = true)]
    [TestFixture]
    public partial class Class1Test
    {
    }
}
