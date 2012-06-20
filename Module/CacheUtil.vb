#Region ".NET Framework Class Import"
#End Region



Public Module CacheUtil

    Public Function GetCache(ByVal ItemName As String) As Object
        Return HttpContext.Current.Cache.Get(ItemName)
    End Function

    Public Sub SetCache(ByVal ItemName As String, ByVal Value As Object, ByVal Duration As TimeSpan)
        'absoluteExpiration must be DateTime.MaxValue or slidingExpiration must be timeSpan.Zero.
        'HttpContext.Current.Cache.Add(ItemName, Value, Nothing, Date.Now.Add(Duration), TimeSpan.FromMinutes(30), Caching.CacheItemPriority.Normal, Nothing)

        HttpContext.Current.Cache.Add(ItemName, Value, Nothing, Date.Now.Add(Duration), TimeSpan.Zero, Caching.CacheItemPriority.Normal, Nothing)
    End Sub

    Public Sub ClearCache(ByVal ItemName As String)
        Try
            HttpContext.Current.Cache.Remove(ItemName)
        Catch ex As Exception
        End Try
    End Sub

End Module
