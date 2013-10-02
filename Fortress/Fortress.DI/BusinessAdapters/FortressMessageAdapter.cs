//-----------------------------------------------------------------------
// <copyright file="FortressMessageAdapter.cs" company="B1C Canada Inc.">
//     Copyright (c) B1C Canada Inc. All rights reserved.
// </copyright>
// <author>Bryan Atkinson</author>
// <email>bryan.atkinson@b1computing.com</email>
//-----------------------------------------------------------------------

namespace Fortress.DI.BusinessAdapters
{
    #region Using Directive(s)
    
    using B1C.SAP.DI.BusinessAdapters;
    using SAPbobsCOM;
    using Helpers;
    using System.Collections.Generic;

    #endregion Using Directive(s)

    public class FortressMessageAdapter : MessageAdapter
    {
        #region Constructor(s)

        /// <summary>
        /// Initializes a new instance of the <see cref="SapObjectAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        public FortressMessageAdapter(Company company) : base(company)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SapObjectAdapter"/> class.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="message">The message</param>
        public FortressMessageAdapter(Company company, Messages message) : base(company, message)
        {
        }

        #endregion Constructor(s)

        #region Static Method(s)

        /// <summary>
        /// Sends the configuration update message.
        /// </summary>
        /// <param name="company">The company.</param>
        /// <param name="orderId">The order id.</param>
        /// <param name="orderLine">The order line.</param>
        /// <param name="newRevision">the new revision number</param>
        public static void SendConfigurationUpdateMessage(Company company, int orderId, string docNumber, int orderLine, string itemCode, int newRevision)
        {
            using (var msg = new FortressMessageAdapter(company))
            {
                IList<string> notificationUsers = FortressConfigurationHelper.GetConfigurationNotificationUser(company);

                // Add each user to the notifications
                foreach (string user in notificationUsers)
                {
                    msg.AddInternalRecipient(user);
                }

                msg.Subject =
                    FortressConfigurationHelper.GetConfigurationNotificationSubject(company);
                msg.Subject = msg.Subject.Replace("{OrderId}", orderId.ToString());
                msg.Subject = msg.Subject.Replace("{LineNum}", orderLine.ToString());
                msg.Subject = msg.Subject.Replace("{oldRev}", (newRevision - 1).ToString());
                msg.Subject = msg.Subject.Replace("{newRev}", newRevision.ToString());
                msg.Subject = msg.Subject.Replace("{ItemCode}", itemCode);
                msg.Subject = msg.Subject.Replace("{DocNumber}", docNumber);

                msg.Body = FortressConfigurationHelper.GetConfigurationNotificationBody(company);
                msg.Body = msg.Body.Replace("{OrderId}", orderId.ToString());
                msg.Body = msg.Body.Replace("{LineNum}", orderLine.ToString());
                msg.Body = msg.Body.Replace("{oldRev}", (newRevision - 1).ToString());
                msg.Body = msg.Body.Replace("{newRev}", newRevision.ToString());
                msg.Body = msg.Body.Replace("{ItemCode}", itemCode);
                msg.Body = msg.Body.Replace("{DocNumber}", docNumber);

                msg.SendMessage();
            }
        }

        #endregion Static Method(s)

    }
}
