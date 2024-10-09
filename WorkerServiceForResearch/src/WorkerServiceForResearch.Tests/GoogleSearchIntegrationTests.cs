using Xunit;
using WorkerServiceForResearch;
using System.Collections.Generic;

public class GoogleSearchIntegrationTests
{
    private readonly GoogleSearch _googleSearch = new GoogleSearch();

    [Fact]
    public void Should_PerformSearchAndReturnResults()
    {
        var searchTerm = "Pokemon";
        var results = _googleSearch.Search(searchTerm);

        Assert.NotEmpty(results);
        Assert.True(results.Count <= 5);
        Assert.All(results, result => Assert.False(string.IsNullOrWhiteSpace(result.Title)));
        Assert.All(results, result => Assert.False(string.IsNullOrWhiteSpace(result.Url)));
    }
}
