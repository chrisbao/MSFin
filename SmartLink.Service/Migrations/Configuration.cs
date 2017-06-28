/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

namespace SmartLink.Service.Migrations
{
    using Entity;
    using System.Data.Entity.Migrations;

    public sealed class Configuration : DbMigrationsConfiguration<SmartLink.Service.SmartlinkDbContext>
    {
        public Configuration()
        {
            AutomaticMigrationsEnabled = true;
            ContextKey = "SmartLink.Service.SmartlinkDbContext";
        }

        protected override void Seed(SmartLink.Service.SmartlinkDbContext context)
        {
            context.SourcePointGroups.AddOrUpdate(x => x.Id,
                    new SourcePointGroup() { Id = 1, Name = "Current year" },
                    new SourcePointGroup() { Id = 2, Name = "Prior year" },
                    new SourcePointGroup() { Id = 3, Name = "PBP" },
                    new SourcePointGroup() { Id = 4, Name = "IC" },
                    new SourcePointGroup() { Id = 5, Name = "MPC" },
                    new SourcePointGroup() { Id = 6, Name = "Revenue" },
                    new SourcePointGroup() { Id = 7, Name = "Gross Margin" },
                    new SourcePointGroup() { Id = 8, Name = "Operating Income" },
                    new SourcePointGroup() { Id = 9, Name = "EPS" },
                    new SourcePointGroup() { Id = 10, Name = "GAAP" },
                    new SourcePointGroup() { Id = 11, Name = "Non-GAAP" },
                    new SourcePointGroup() { Id = 12, Name = "Outlook" },
                    new SourcePointGroup() { Id = 13, Name = "Momentum Statement" }
                    );

            context.CustomFormats.AddOrUpdate(x => x.Id,
                    new CustomFormat() { Id = 1, Name = "ConvertToHundreds", DisplayName = "Convert to hundreds", Description = "Divide source point by 100 and insert 0 and decimal" },
                    new CustomFormat() { Id = 2, Name = "ConvertToThousands", DisplayName = "Convert to thousands", Description = "Divide source point by 1,000" },
                    new CustomFormat() { Id = 3, Name = "ConvertToMillions", DisplayName = "Convert to millions", Description = "Divide source point by 1,000,000" },
                    new CustomFormat() { Id = 4, Name = "ConvertToBillions", DisplayName = "Convert to billions", Description = "Divide source point by 1,000,000,000" },
                    new CustomFormat() { Id = 5, Name = "AddDecimalPlace", DisplayName = "Add decimal place", Description = "Display additional decimal place" },
                    new CustomFormat() { Id = 6, Name = "ShowNegativesAsPositives", DisplayName = "Show negatives as positives", Description = "Multiply by -1" },
                    new CustomFormat() { Id = 7, Name = "IncludeThousandDescriptor", DisplayName = "Include \"thousand\" descriptor", Description = "Insert thousand after numerical value" },
                    new CustomFormat() { Id = 8, Name = "IncludeMillionDescriptor", DisplayName = "Include \"million\" descriptor", Description = "Insert million after numerical value" },
                    new CustomFormat() { Id = 9, Name = "IncludeBillionDescriptor", DisplayName = "Include \"billion\" descriptor", Description = "Insert billion after numerical value" },
                    new CustomFormat() { Id = 10, Name = "IncludeDollarSymbol", DisplayName = "Include $ symbol", Description = "Add dollar sign to front of source point value" },
                    new CustomFormat() { Id = 11, Name = "ExcludeDollarSymbol", DisplayName = "Exclude $ symbol", Description = "Remove dollar sign to front of source point value" },
                    new CustomFormat() { Id = 12, Name = "DateShowLongDateFormat", DisplayName = "Date: Show long date format", Description = "Convert MM/DD/YYYY to Month DD, YYYY" },
                    new CustomFormat() { Id = 13, Name = "DateShowYearOnly", DisplayName = "Date: Show year only", Description = "Convert MM/DD/YYYY to YYYY" },
                    new CustomFormat() { Id = 14, Name = "ConvertNegativeSymbolToParenthesis", DisplayName = "Convert negative symbol to parenthesis", Description = "Remove '-' symbol and replace with '( )'" }
                    );
        }
    }
}