/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using System;
using System.Windows.Forms;

namespace ContosoO365DocSync.EncryptTool
{
    public partial class Main : Form
    {
        private EncryptionService _encryptService;
        public Main()
        {
            _encryptService = new EncryptionService();
            InitializeComponent();
        }

        /// <summary>
        /// Encrypt string handle.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnEncryptString_Click(object sender, EventArgs e)
        {
            try
            {
                txtTarget.Text = _encryptService.EncryptString(txtSource.Text);
            }
            catch (Exception ex)
            {

            }
        }

        /// <summary>
        /// Decrypt string handle.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDecryptString_Click(object sender, EventArgs e)
        {
            try
            {
                txtTarget.Text = _encryptService.DecryptString(txtSource.Text);
            }
            catch (Exception ex)
            {

            }
        }
    }
}
