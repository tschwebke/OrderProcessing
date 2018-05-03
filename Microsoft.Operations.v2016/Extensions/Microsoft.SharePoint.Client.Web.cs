using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;

public static class SharePointExtensions
{
    public static Folder GetFolder(this Web web, string fullFolderUrl)
    {
        if (string.IsNullOrEmpty(fullFolderUrl))
            throw new ArgumentNullException("fullFolderUrl");

        if (!web.IsPropertyAvailable("ServerRelativeUrl"))
        {
            web.Context.Load(web, w => w.ServerRelativeUrl);
            web.Context.ExecuteQuery();
        }
        var folder = web.GetFolderByServerRelativeUrl(web.ServerRelativeUrl + fullFolderUrl);
        web.Context.Load(folder);
        web.Context.ExecuteQuery();
        return folder;
    }

    public static Folder GetFolderCreateIfNotExists(this Web web, Folder parent, string nameOfDesiredFolder)
    {
        Folder folderFinal = null;

        if (!web.IsPropertyAvailable("ServerRelativeUrl"))
        {
            web.Context.Load(web, w => w.ServerRelativeUrl);
            web.Context.ExecuteQuery();
        }

        try
        {
            folderFinal = web.GetFolderByServerRelativeUrl(parent.ServerRelativeUrl + '/' + nameOfDesiredFolder);
            web.Context.Load(folderFinal);
            web.Context.ExecuteQuery();
        }
        catch
        {
            // doesn't exist, so try and create it.
            folderFinal = parent.Folders.Add(nameOfDesiredFolder);
            web.Context.Load(folderFinal);
            web.Context.ExecuteQuery();
        }

        return folderFinal;
    }

    public static void LoadContent(this Web web, out Dictionary<string, IEnumerable<Folder>> listsFolders, out Dictionary<string, IEnumerable<File>> listsFiles)
    {
        listsFolders = new Dictionary<string, IEnumerable<Folder>>();
        listsFiles = new Dictionary<string, IEnumerable<File>>();
        var listsItems = new Dictionary<string, IEnumerable<ListItem>>();

        var ctx = web.Context;
        var lists = ctx.LoadQuery(web.Lists.Where(l => l.BaseType == BaseType.DocumentLibrary));
        ctx.ExecuteQuery();

        foreach (var list in lists)
        {
            var items = list.GetItems(CamlQuery.CreateAllItemsQuery());
            ctx.Load(items, icol => icol.Include(i => i.FileSystemObjectType, i => i.File, i => i.Folder));
            listsItems[list.Title] = items;
        }
        ctx.ExecuteQuery();

        foreach (var listItems in listsItems)
        {
            listsFiles[listItems.Key] = listItems.Value.Where(i => i.FileSystemObjectType == FileSystemObjectType.File).Select(i => i.File);
            listsFolders[listItems.Key] = listItems.Value.Where(i => i.FileSystemObjectType == FileSystemObjectType.Folder).Select(i => i.Folder);
        }
    }
}