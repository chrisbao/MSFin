/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using System;
using System.Collections.Generic;

namespace ContosoO365DocSync.Service
{
    public class Comparer<T> : IEqualityComparer<T>
    {
        private readonly Func<T, T, bool> _equalityComparer;

        public Comparer(Func<T, T, bool> equalityComparer)
        {
            _equalityComparer = equalityComparer;
        }

        public bool Equals(T first, T second)
        {
            return _equalityComparer(first, second);
        }

        public int GetHashCode(T value)
        {
            return value.GetHashCode();
        }
    }
}