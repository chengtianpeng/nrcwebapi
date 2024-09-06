
FROM mcr.microsoft.com/dotnet/aspnet:7.0 AS base
WORKDIR /app
EXPOSE 80

FROM mcr.microsoft.com/dotnet/sdk:7.0 AS build
WORKDIR /src
COPY ["NRCWebApi/NRCWebApi.csproj", "NRCWebApi/"]
RUN dotnet restore "NRCWebApi/NRCWebApi.csproj"
COPY . .
WORKDIR "/src/NRCWebApi"
RUN dotnet build "NRCWebApi.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "NRCWebApi.csproj" -c Release -o /app/publish 

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "NRCWebApi.dll"]


