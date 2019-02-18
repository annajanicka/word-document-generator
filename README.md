# word-document-generator

This package allows you to generate \*.docx document based on the \*.docx template and provided model.
It supports type formatting and tables.

## Example
### \*.doc content

Test document

{model.Property1}

{model.Property2}

{model.Property3}

{model.Property4:F}

{model.Property5:G}

{model.Property6:D}

{model.Property7:hh}

{model.Property8}

{model.Property9:N}

{model.SubModel.Property1}

{model.SubModel.Property2}


<table>
    <tbody>
        <tr>
            <td rowspan=2>{i}</td>
            <td>{model.Items[i].Property1} </td>
            <td>{model.Items[i].Property2}</td>
        </tr>
        <tr>
            <td colspan=2>{i-related} related content</td>
        </tr>
    </tbody>
</table>

### \*.doc output

Test document

1eed3f08-7a44-4437-bcf9-1cfe3e77c28d

e5cecf0e-7ca7-4709-b56d-21e205adf567

123123

2.20

18/02/2019 10:21:11 AM

Monday, 18 February 2019

00

<0, 1>

6e0e30c971e14e179bfa60f86b8648ac

cd7138aa-6c67-4af8-bc43-59da2198ce63

96b34977-94b1-4e0f-9e71-f4123da6c754

<table>
    <tbody>
        <tr>
            <td rowspan=2>1</td>
            <td>16c69507-f0da-451c-a783-93cbd3a3793a</td>
            <td>fe569a05-0721-46f2-a970-e10beb8d4532</td>
        </tr>
        <tr>
            <td colspan=2>related content</td>
        </tr>
   	<tr>
            <td rowspan=2>2</td>
            <td>e45494a8-1268-4bec-b5f9-929c014c9b74</td>
            <td>b6940a2b-6755-4a66-90d6-699ea6cdef2f</td>
        </tr>
        <tr>
            <td colspan=2>related content</td>
        </tr>
    	<tr>
            <td rowspan=2>3</td>
            <td>cc535fe0-1d4b-42b3-937f-322b493ac207</td>
            <td>f8e83136-d21a-4563-8914-fce0718b9bbe</td>
        </tr>
    </tbody>
</table>

### code
 ```csharp
var assembly = Assembly.GetExecutingAssembly();
using (var stream = assembly.GetManifestResourceStream("DocumentGenerator.Example.form.docx"))
using (MemoryStream memoryStream = new MemoryStream())
{
    stream.CopyTo(memoryStream);
    var result = new WordDocumentGenerator().GetDocument(GetModel(), memoryStream.ToArray());
    {
        File.WriteAllBytes("output.docx", result);
    }
}

```
Model preparation
 ```csharp
private static Model GetModel()
{
    return new Model(
        Guid.NewGuid().ToString(),
        Guid.NewGuid().ToString(),
        123123,
        2.2M,
        DateTime.Now,
        DateTimeOffset.Now,
        TimeSpan.FromDays(123232),
        Vector2.UnitY,
        Guid.NewGuid(),
        new SubModel(Guid.NewGuid().ToString(), Guid.NewGuid().ToString()),
        new[]
        {
            new Item(Guid.NewGuid().ToString(), Guid.NewGuid().ToString()),
            new Item(Guid.NewGuid().ToString(), Guid.NewGuid().ToString()),
            new Item(Guid.NewGuid().ToString(), Guid.NewGuid().ToString())
        });
}
```

  ```csharp       
public class Model
{
    public string Property1 { get; }
    public string Property2 { get; }
    public int Property3 { get; }
    public decimal Property4 { get; }
    public DateTime Property5 { get; }
    public DateTimeOffset Property6 { get; }
    public TimeSpan Property7 { get; }
    public Vector2 Property8 { get; }
    public Guid Property9 { get; }
    public SubModel SubModel { get; }
    public Item[] Items { get; }

    public Model(string property1,
        string property2,
        int property3,
        decimal property4,
        DateTime property5,
        DateTimeOffset property6,
        TimeSpan property7,
        Vector2 property8,
        Guid property9,
        SubModel subModel,
        Item[] items)
    {
        Property1 = property1;
        Property2 = property2;
        Property3 = property3;
        Property4 = property4;
        Property5 = property5;
        Property6 = property6;
        Property7 = property7;
        Property8 = property8;
        Property9 = property9;
        SubModel = subModel;
        Items = items;
    }    
}

public class SubModel
{
    public string Property1 { get; }
    public string Property2 { get; }

    public SubModel(string property1, string property2)
    {
        Property1 = property1;
        Property2 = property2;
    }
}

public class Item
{
    public string Property1 { get; }
    public string Property2 { get; }

    public Item(string property1, string property2)
    {
        Property1 = property1;
        Property2 = property2;
    }
}
```
