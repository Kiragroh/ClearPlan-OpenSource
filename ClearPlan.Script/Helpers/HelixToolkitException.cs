// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ClearPlanException.cs" company="Helix Toolkit">
//   Copyright (c) 2014 Helix Toolkit contributors
// </copyright>
// <summary>
//   Represents errors that occurs in the Helix 3D Toolkit.
// </summary>
// --------------------------------------------------------------------------------------------------------------------
#if SHARPDX
#if NETFX_CORE
namespace ClearPlan.UWP
#else
namespace ClearPlan.SharpDX
#endif
#else
namespace ClearPlan
#endif
{
    using System;

#pragma warning disable 0436
    /// <summary>
    /// Represents errors that occurs in the Helix 3D Toolkit.
    /// </summary>
#if !NETFX_CORE
    [Serializable]
#endif
    public class ClearPlanException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ClearPlanException"/> class.
        /// </summary>
        /// <param name="formatString">
        /// The format string.
        /// </param>
        /// <param name="args">
        /// The args.
        /// </param>
        public ClearPlanException(string formatString, params object[] args)
            : base(string.Format(formatString, args))
        {
        }
    }
#pragma warning restore 0436
}
