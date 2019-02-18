using System;
using NUnit.Framework;
using Shouldly;

namespace DocumentGenerator.Tests
{
    [TestFixture]
    public class WhenParseModel
    {
        private FakeModel _model;
        private ModelParser _modelParser;

        [SetUp]
        public void Setup()
        {
            _model = new FakeModel(
                "Value of property 1",
                "Value of property 2",
                new FakeModel.FakeSubModel("Value of property 1 of sub model", "Value of property 2 of sub model"),
                new[]
                {
                    new FakeModel.FakeItem("Value of property 1 of fake item 1", "Value of property 2 of fake item 1"),
                    new FakeModel.FakeItem("Value of property 1 of fake item 1", "Value of property 2 of fake item 1")
                });

            _modelParser = new ModelParser();
        }

        [Test]
        public void WhenPropertiesExistReturnValues()
        {
            _modelParser.GetPropertyValue(_model, "Property1").ShouldBe(_model.Property1);
            _modelParser.GetPropertyValue(_model, "Property2").ShouldBe(_model.Property2);
            _modelParser.GetPropertyValue(_model, "SubModel.Property1").ShouldBe(_model.SubModel.Property1);
            _modelParser.GetPropertyValue(_model, "SubModel.Property2").ShouldBe(_model.SubModel.Property2);
            _modelParser.GetPropertyValue(_model, "Items[0].Property1").ShouldBe(_model.Items[0].Property1);
            _modelParser.GetPropertyValue(_model, "Items[0].Property2").ShouldBe(_model.Items[0].Property2);
            _modelParser.GetPropertyValue(_model, "Items[1].Property1").ShouldBe(_model.Items[1].Property1);
            _modelParser.GetPropertyValue(_model, "Items[1].Property2").ShouldBe(_model.Items[1].Property2);
        }

        [Test]
        public void WhenPropertyDoesNotExistThrowException()
        {
            var exception = Assert.Throws<NullReferenceException>(() => _modelParser.GetPropertyValue(_model, "NonExisitngProperty"));
            Assert.That(exception.Message, Is.EqualTo("property is null for 'NonExisitngProperty'. Document template contains unknown model expression."));
        }
    }
}
