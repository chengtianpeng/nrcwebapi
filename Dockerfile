#See https://aka.ms/customizecontainer to learn how to customize your debug container and how Visual Studio uses this Dockerfile to build your images for faster debugging.

FROM mcr.microsoft.com/dotnet/aspnet:7.0 AS base
WORKDIR /app
EXPOSE 80
EXPOSE 443
#
#FROM mcr.microsoft.com/dotnet/sdk:7.0 AS build
#WORKDIR /src
#COPY ["NRCWebApi/NRCWebApi.csproj", "NRCWebApi/"]
#RUN dotnet restore "NRCWebApi/NRCWebApi.csproj"
#COPY . .
#WORKDIR "/src/NRCWebApi"
#RUN dotnet build "NRCWebApi.csproj" -c Release -o /app/build
#
#FROM build AS publish
#RUN dotnet publish "NRCWebApi.csproj" -c Release -o /app/publish /p:UseAppHost=false
#
#FROM base AS final
#WORKDIR /app
#COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "NRCWebApi.dll"]