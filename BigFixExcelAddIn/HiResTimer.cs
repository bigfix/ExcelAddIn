/*
class HiResTimer

    by: Shawn A. Van Ness
    rev: 2001.10.30

This simple little C# class wraps Kernel32's QueryPerformanceCounter API. 
Usage is straightforward:

    HiResTimer hrt = new HiResTimer();

    hrt.Start();
    DoSomethingLengthy();
    hrt.Stop();

    Console.WriteLine("{0}", hrt.ElapsedMicroseconds);

The hires timer API deals in unsigned, 64-bit (ulong) values, so the problem 
of rollover (or "lapping") is not an issue -- even on a box with a 1GHz timer 
frequency, 500 years will pass by before the counter laps its start position 
(and 500 years is probably far longer than the MTBF of a sizzling hot 1GHz 
Pentium).

Not all systems support the hires timer API (pre-Pentium II? non-Intel?). The 
constructor will throw a Win32Exception on those systems.

Because the code must P/Invoke down to Kernel32.dll, it may not be desirable 
to include in retail code.  But because it's so awkward to comment-out all of 
the references to HiResTimer (4 non-contiguous lines, in the above example!) 
I've included a handy preprocessor symbol (NOTIMER) to render the class inert.
*/

using System;
using System.Runtime.InteropServices; // for DllImport attribute
using System.ComponentModel; // for Win32Exception class
using System.Threading; // for Thread.Sleep method

class HiResTimer
{
    // Construction

    public HiResTimer()
    {
#if (!NOTIMER)
        a = b = 0UL;
        if (QueryPerformanceFrequency(out f) == 0)
            throw new Win32Exception();
#endif
    }

    // Properties

    public ulong ElapsedTicks
    {
#if (!NOTIMER)
        get
        { return (b - a); }
#else
        get
        { return 0UL; }
#endif
    }

    public ulong ElapsedMicroseconds
    {
#if (!NOTIMER)
        get
        {
            ulong d = (b - a);
            if (d < 0x10c6f7a0b5edUL) // 2^64 / 1e6
                return (d * 1000000UL) / f;
            else
                return (d / f) * 1000000UL;
        }
#else
        get
        { return 0UL; }
#endif
    }

    public TimeSpan ElapsedTimeSpan
    {
#if (!NOTIMER)
        get
        {
            ulong t = 10UL * ElapsedMicroseconds;
            if ((t & 0x8000000000000000UL) == 0UL)
                return new TimeSpan((long)t);
            else
                return TimeSpan.MaxValue;
        }
#else
        get
        { return TimeSpan.Zero; }
#endif
    }

    public ulong Frequency
    {
#if (!NOTIMER)
        get
        { return f; }
#else
        get
        { return 1UL; }
#endif
    }

    // Methods

    public void Start()
    {
#if (!NOTIMER)
        Thread.Sleep(0);
        QueryPerformanceCounter(out a);
#endif
    }

    public ulong Stop()
    {
#if (!NOTIMER)
        QueryPerformanceCounter(out b);
        return ElapsedTicks;
#else
        return 0UL;
#endif
    }

    // Implementation

#if (!NOTIMER)
    [DllImport("kernel32.dll", SetLastError = true)]
    protected static extern
        int QueryPerformanceFrequency(out ulong x);

    [DllImport("kernel32.dll")]
    protected static extern
        int QueryPerformanceCounter(out ulong x);

    protected ulong a, b, f;
#endif
}
