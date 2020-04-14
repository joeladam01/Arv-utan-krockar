using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xna.Framework;
using Microsoft.Xna.Framework.Graphics;
using Microsoft.Xna.Framework.Input;
using OfficeOpenXml;
using System.IO;

namespace Spaceshooter
{
    public class Game1 : Game
    {
        GraphicsDeviceManager graphics;
        SpriteBatch spriteBatch;
        List<Enemy> enemies;
        Texture2D goldCoinSprite;
        static ExcelWorksheet sheet;
        static ExcelPackage package;
        static FileInfo file;
        int raknare;



        public Game1()
        {
            graphics = new GraphicsDeviceManager(this);
            Content.RootDirectory = "Content";
        }

        protected override void Initialize()
        {
            
            base.Initialize();
            graphics.PreferredBackBufferWidth = 1500;
            graphics.PreferredBackBufferHeight = 1500;
            graphics.ApplyChanges();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            file = new FileInfo(@"C:\Users\Joel\Desktop\ArvUtanKrockar.xlsx");

            package = new ExcelPackage(file);

            sheet = package.Workbook.Worksheets.Add("Blad1");

            raknare = 1;

            sheet.Cells[$"A{raknare}"].Value = "Update";
            sheet.Cells[$"B{raknare++}"].Value = "Draw";

        }

        protected override void LoadContent()
        {
            spriteBatch = new SpriteBatch(GraphicsDevice);
           

            // Skapa fiender
            enemies = new List<Enemy>();
            Random random = new Random();
            Texture2D tmpSprite = Content.Load<Texture2D>("images/enemies/mine");
            int posX = 0;
            int posY = 0;
            for (int i = 0; i < 100; i++)
            {
                if (posX < 1350)
                {
                    posX += 150;
                }
                else
                {
                    posX = 75;
                    posY += 150;
                }
                Mine temp = new Mine(tmpSprite, posX, posY);
                    
                // Lägg till i listan
                enemies.Add(temp);
            }

            tmpSprite = Content.Load<Texture2D>("images/enemies/tripod");
            int posX1 = 75;
            int posY1 = 75;
            for (int i = 0; i < 100; i++)
            {
                if (posX1 < 1350)
                {
                    posX1 += 150;
                }
                else
                {
                    posX1 = 75;
                    posY1 += 150;
                }
                Tripod temp = new Tripod(tmpSprite, posX1, posY1);

                // Lägg till i listan
                enemies.Add(temp);
            }

            
            goldCoinSprite = Content.Load<Texture2D>("images/powerups/coin");
        }

        protected override void UnloadContent()
        {
        }
        
        protected override void Update(GameTime gameTime)
        {
            var watch = System.Diagnostics.Stopwatch.StartNew();
            if (GamePad.GetState(PlayerIndex.One).Buttons.Back == ButtonState.Pressed)
                this.Exit();

           

            // Gå igenom alla fiender
            foreach (Enemy e in enemies.ToList())
            {
                e.Update(Window);
            }

            base.Update(gameTime);
            watch.Stop();
            var elapsedMS = watch.ElapsedTicks;
            if (raknare <= 1200)
            {
                sheet.Cells[$"A{raknare}"].Value = elapsedMS;
            }
            else
            {
                package.Save();
            }


        }

        protected override void Draw(GameTime gameTime)
        {
            var watch = System.Diagnostics.Stopwatch.StartNew();
            GraphicsDevice.Clear(Color.CornflowerBlue);
            spriteBatch.Begin();

            foreach (Enemy e in enemies)
                e.Draw(spriteBatch);

            spriteBatch.End();
            base.Draw(gameTime);

            watch.Stop();
            var elapsedMS = watch.ElapsedTicks;
            if (raknare <= 1200)
            {
                sheet.Cells[$"B{raknare++}"].Value = elapsedMS;
            }
        }
    }
}
