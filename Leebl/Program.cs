using Leebl.ExchangeWebServices;
using Leebl.Objects;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Web.Services.Protocols;

namespace Leebl
{
  internal class Program
  {
    private static void Main()
    {
      Console.WriteLine("");
      Console.WriteLine("// Leebl 1.0 //////////////////");
      Console.WriteLine("// Copyright Coen Software B.V.");
      Console.WriteLine("");

      Console.Write("Exchange username: ");
      var username = Console.ReadLine();

      Console.Write("Exchange password: ");
      var password = Console.ReadLine();

      Console.Write("Exchange folder name: ");
      var folderName = Console.ReadLine();

      try
      {
        Console.WriteLine();
        Console.WriteLine("Processing...");
        Console.WriteLine();

        Process(username, password, folderName);

        Console.WriteLine();
        Console.WriteLine("Finished!");
      }
      catch (WebException ex)
      {
        Console.WriteLine("Something went wrong when contacting Exchange - are your credentials correct?");
        Console.WriteLine();
        Console.WriteLine(ex);
      }
      catch (SoapException ex)
      {
        Console.WriteLine("Something went wrong when processing Exchange - is the folder name correct?");
        Console.WriteLine();
        Console.WriteLine(ex);
      }
      catch (Exception ex)
      {
        Console.WriteLine("An error occurred, please try again later.");
        Console.WriteLine();
        Console.WriteLine(ex);
      }
      finally
      {
        Console.WriteLine();
        Console.WriteLine("<Press Enter to exit>");
        Console.ReadLine();
      }
    }

    private static void Process(string username, string password, string folderName)
    {
      Console.WriteLine(" > Opening file...");

      var path = Directory.GetCurrentDirectory() + @"\Leebl.csv";
      var leads = new LeadCollection(path);

      var binding = new ExchangeServiceBinding
      {
        Url = "https://amxprd0510.outlook.com/ews/exchange.asmx",
        Credentials = new NetworkCredential(username, password),
        RequestServerVersionValue = new RequestServerVersion { Version = ExchangeVersionType.Exchange2010 }
      };

      #region Get folder

      Console.WriteLine(" > Retrieving folder...");

      var folderRequest = new FindFolderType
      {
        Traversal = FolderQueryTraversalType.Deep,
        FolderShape = new FolderResponseShapeType { BaseShape = DefaultShapeNamesType.IdOnly },
        ParentFolderIds = new BaseFolderIdType[]
        {
          new DistinguishedFolderIdType { Id = DistinguishedFolderIdNameType.root }
        },
        Restriction = new RestrictionType
        {
          Item = new ContainsExpressionType
          {
            ContainmentMode = ContainmentModeType.Substring,
            ContainmentModeSpecified = true,
            ContainmentComparison = ContainmentComparisonType.IgnoreCase,
            ContainmentComparisonSpecified = true,
            Item = new PathToUnindexedFieldType
            {
              FieldURI = UnindexedFieldURIType.folderDisplayName
            },
            Constant = new ConstantValueType
            {
              Value = folderName
            }
          }
        }
      };

      var folderResponse = binding.FindFolder(folderRequest);
      var folderIds = new List<BaseFolderIdType>();

      foreach (var folder in folderResponse.ResponseMessages.Items
        .Select(x => (x as FindFolderResponseMessageType))
        .Where(x => x != null))
      {
        folderIds.AddRange(folder.RootFolder.Folders.Select(y => y.FolderId));
      }

      #endregion

      #region Get items

      Console.WriteLine(" > Retrieving items...");

      var itemRequest = new FindItemType
      {
        Traversal = ItemQueryTraversalType.Shallow,
        ItemShape = new ItemResponseShapeType { BaseShape = DefaultShapeNamesType.Default },
        ParentFolderIds = folderIds.ToArray()
      };

      var itemResponse = binding.FindItem(itemRequest);
      var itemIds = new List<BaseItemIdType>();

      foreach (var item in itemResponse.ResponseMessages.Items
        .Select(x => (x as FindItemResponseMessageType))
        .Where(x => x != null)
        .Where(x => x.RootFolder != null && x.RootFolder.TotalItemsInView > 0))
      {
        itemIds.AddRange(((ArrayOfRealItemsType)item.RootFolder.Item).Items.Select(y => y.ItemId));
      }

      #endregion

      #region Get bodies

      Console.WriteLine(" > Parsing " + itemIds.Count + " messages...");

      var messageRequest = new GetItemType
      {
        ItemShape = new ItemResponseShapeType { BaseShape = DefaultShapeNamesType.AllProperties },
        ItemIds = itemIds.ToArray()
      };

      var messageResponse = binding.GetItem(messageRequest);

      foreach (var message in messageResponse.ResponseMessages.Items
        .Select(x => (x as ItemInfoResponseMessageType))
        .Where(x => x != null)
        .Select(x => x.Items.Items[0])
        .Where(x => x != null))
      {
        leads.Add(message.Body.Value, message.DateTimeSent);
      }

      #endregion

      Console.WriteLine(" > Saving to file...");
      leads.Save(path);
    }
  }
}