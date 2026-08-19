#if NET40_TEST_RUNNER
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Xunit
{
    [AttributeUsage(AttributeTargets.Method)]
    internal sealed class FactAttribute : Attribute { }

    [AttributeUsage(AttributeTargets.Method)]
    internal sealed class TheoryAttribute : Attribute { }

    [AttributeUsage(AttributeTargets.Method, AllowMultiple = true)]
    internal sealed class InlineDataAttribute : Attribute
    {
        internal InlineDataAttribute(params object[] data) { Data = data; }
        internal object[] Data { get; private set; }
    }

    internal static class Assert
    {
        internal static void True(bool condition, string? message = null)
        {
            if (!condition) throw new Exception(message ?? "Assert.True failed.");
        }

        internal static void Equal<T>(T expected, T actual)
        {
            var expectedItems = expected as IEnumerable;
            var actualItems = actual as IEnumerable;
            if (!(expected is string) && expectedItems != null && actualItems != null)
            {
                var left = expectedItems.Cast<object>().ToArray();
                var right = actualItems.Cast<object>().ToArray();
                if (left.Length == right.Length && left.SequenceEqual(right)) return;
            }
            else if (object.Equals(expected, actual))
            {
                return;
            }

            throw new Exception(string.Format("Assert.Equal failed. Expected: {0}; actual: {1}.", expected, actual));
        }

        internal static void Empty(IEnumerable values)
        {
            if (values.Cast<object>().Any()) throw new Exception("Assert.Empty failed.");
        }

        internal static void Single(IEnumerable values)
        {
            if (values.Cast<object>().Count() != 1) throw new Exception("Assert.Single failed.");
        }

        internal static void All<T>(IEnumerable<T> values, Action<T> assertion)
        {
            foreach (var value in values) assertion(value);
        }

        internal static void Contains(string expectedSubstring, string actual)
        {
            if (actual == null || actual.IndexOf(expectedSubstring, StringComparison.Ordinal) < 0)
                throw new Exception("Assert.Contains failed.");
        }

        internal static T Throws<T>(Action action) where T : Exception
        {
            try { action(); }
            catch (T exception) { return exception; }
            catch (Exception exception)
            {
                throw new Exception("Assert.Throws received " + exception.GetType().FullName + " instead of " + typeof(T).FullName + ".", exception);
            }

            throw new Exception("Assert.Throws did not receive " + typeof(T).FullName + ".");
        }
    }
}

namespace XCelReportEngine.Tests
{
    internal static class Net40TestRunner
    {
        public static int Main()
        {
            var failures = new List<string>();
            var executed = 0;
            var testType = typeof(ReportEngineApiTests);
            var instance = Activator.CreateInstance(testType);

            foreach (var method in testType.GetMethods(BindingFlags.Instance | BindingFlags.Public))
            {
                var fact = method.GetCustomAttributes(typeof(Xunit.FactAttribute), false).Any();
                var theory = method.GetCustomAttributes(typeof(Xunit.TheoryAttribute), false).Any();
                if (!fact && !theory) continue;

                var cases = theory
                    ? method.GetCustomAttributes(typeof(Xunit.InlineDataAttribute), false).Cast<Xunit.InlineDataAttribute>().Select(item => item.Data)
                    : new[] { new object[0] }.AsEnumerable();

                foreach (var arguments in cases)
                {
                    executed++;
                    try
                    {
                        method.Invoke(instance, arguments);
                    }
                    catch (TargetInvocationException exception)
                    {
                        failures.Add(method.Name + ": " + (exception.InnerException ?? exception));
                    }
                    catch (Exception exception)
                    {
                        failures.Add(method.Name + ": " + exception);
                    }
                }
            }

            Console.WriteLine(".NET Framework 4.0 tests: {0} executed, {1} failed.", executed, failures.Count);
            foreach (var failure in failures) Console.Error.WriteLine(failure);
            return failures.Count == 0 ? 0 : 1;
        }
    }
}
#endif
