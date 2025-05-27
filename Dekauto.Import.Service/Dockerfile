# See https://aka.ms/customizecontainer to learn how to customize your debug container and how Visual Studio uses this Dockerfile to build your images for faster debugging.

# This stage is used when running from VS in fast mode (Default for Debug configuration)
FROM mcr.microsoft.com/dotnet/aspnet:8.0-alpine AS base
# Предустановка curl внутри линукса
USER root
RUN apk update && apk add --no-cache curl
USER $APP_UID
WORKDIR /app

# HTTP и HTTPS порты
EXPOSE 5503 5504  

# This stage is used to build the service project
FROM mcr.microsoft.com/dotnet/sdk:8.0-alpine AS build
ARG BUILD_CONFIGURATION=Release
WORKDIR /src
COPY ["Dekauto.Import.Service.csproj", "."]
RUN dotnet restore "./Dekauto.Import.Service.csproj"
COPY . .
WORKDIR "/src/."
RUN dotnet build "./Dekauto.Import.Service.csproj" -c $BUILD_CONFIGURATION -o /app/build

# This stage is used to publish the service project to be copied to the final stage
FROM build AS publish
ARG BUILD_CONFIGURATION=Release
RUN dotnet publish "./Dekauto.Import.Service.csproj" -c $BUILD_CONFIGURATION -o /app/publish /p:UseAppHost=false

# This stage is used in production or when running from VS in regular mode (Default when not using the Debug configuration)
FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "Dekauto.Import.Service.dll"]